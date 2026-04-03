#!/usr/bin/env python3
"""
add_image.py — Add an image to an unpacked PPTX slide.

Usage:
    python add_image.py UNPACKED_DIR SLIDE_NUM IMAGE_PATH [OPTIONS]

Options:
    --x EMU          X position (default: 457200 = 0.5")
    --y EMU          Y position (default: 1200000)
    --cx EMU         Width (default: auto from image aspect ratio)
    --cy EMU         Height (default: auto from image aspect ratio)
    --max-cx EMU     Max width constraint (default: 8229600 = full content width)
    --max-cy EMU     Max height constraint (default: 3200000)
    --name NAME      Shape name (default: "Image")
    --round EMU      Corner radius for rounded rectangle crop (0 = no rounding)

Copies the image into ppt/media/, updates Content_Types.xml,
creates the relationship in slide .rels, and outputs the <p:pic>
XML snippet to paste into the slide XML.

Example:
    python add_image.py unpacked/ 2 photo.jpg --x 457200 --y 1200000 --max-cx 4000000
"""

import argparse
import os
import re
import shutil
import struct
import sys
from pathlib import Path


# --- Image dimension detection (no PIL dependency) ---

def get_png_dimensions(path):
    with open(path, 'rb') as f:
        f.read(8)  # skip signature
        f.read(4)  # chunk length
        f.read(4)  # chunk type (IHDR)
        w = struct.unpack('>I', f.read(4))[0]
        h = struct.unpack('>I', f.read(4))[0]
    return w, h


def get_jpeg_dimensions(path):
    with open(path, 'rb') as f:
        f.read(2)  # SOI
        while True:
            marker = f.read(2)
            if len(marker) < 2:
                break
            if marker[0] != 0xFF:
                break
            mtype = marker[1]
            if mtype in (0xC0, 0xC1, 0xC2):
                f.read(3)  # length + precision
                h = struct.unpack('>H', f.read(2))[0]
                w = struct.unpack('>H', f.read(2))[0]
                return w, h
            else:
                length = struct.unpack('>H', f.read(2))[0]
                f.read(length - 2)
    return None, None


def get_image_dimensions(path):
    """Return (width, height) in pixels without PIL."""
    ext = Path(path).suffix.lower()
    if ext == '.png':
        return get_png_dimensions(path)
    elif ext in ('.jpg', '.jpeg'):
        return get_jpeg_dimensions(path)
    else:
        # Try PIL as fallback
        try:
            from PIL import Image
            with Image.open(path) as img:
                return img.size
        except ImportError:
            return None, None


# --- Content type mapping ---

CONTENT_TYPES = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.webp': 'image/webp',
    '.bmp': 'image/bmp',
    '.tiff': 'image/tiff',
    '.tif': 'image/tiff',
}


def find_next_media_index(media_dir):
    """Find next available imageN index in ppt/media/."""
    existing = []
    if media_dir.exists():
        for f in media_dir.iterdir():
            m = re.match(r'image(\d+)', f.stem)
            if m:
                existing.append(int(m.group(1)))
    return max(existing, default=0) + 1


def find_next_rid(rels_path):
    """Find next available rId in a .rels file."""
    if not rels_path.exists():
        return 'rId1', None
    content = rels_path.read_text(encoding='utf-8')
    ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', content)]
    next_id = max(ids, default=0) + 1
    return f'rId{next_id}', content


def add_content_type(ct_path, ext):
    """Add image extension to [Content_Types].xml if not present."""
    content = ct_path.read_text(encoding='utf-8')
    ext_no_dot = ext.lstrip('.')
    # Check if already declared
    if re.search(rf'Extension="{ext_no_dot}"', content, re.IGNORECASE):
        return
    ct = CONTENT_TYPES.get(ext.lower())
    if not ct:
        print(f"WARNING: Unknown content type for {ext}", file=sys.stderr)
        return
    # Insert before closing </Types>
    insert = f'<Default Extension="{ext_no_dot}" ContentType="{ct}"/>'
    content = content.replace('</Types>', f'{insert}\n</Types>')
    ct_path.write_text(content, encoding='utf-8')
    print(f"  Added Content_Type: {ext_no_dot} -> {ct}")


def add_relationship(rels_path, rid, target):
    """Add image relationship to slide .rels file."""
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    new_rel = f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/>'

    if rels_path.exists():
        content = rels_path.read_text(encoding='utf-8')
        content = content.replace('</Relationships>', f'{new_rel}\n</Relationships>')
    else:
        content = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            f'{new_rel}\n'
            '</Relationships>'
        )
    rels_path.parent.mkdir(parents=True, exist_ok=True)
    rels_path.write_text(content, encoding='utf-8')
    print(f"  Added relationship: {rid} -> {target}")


def calculate_dimensions(img_w, img_h, max_cx, max_cy, cx_override, cy_override):
    """Calculate EMU dimensions preserving aspect ratio."""
    if cx_override and cy_override:
        return cx_override, cy_override

    # Convert pixels to EMU (96 DPI assumed: 1px = 914400/96 = 9525 EMU)
    px_to_emu = 9525
    native_cx = img_w * px_to_emu
    native_cy = img_h * px_to_emu

    if cx_override:
        scale = cx_override / native_cx
        return cx_override, int(native_cy * scale)
    if cy_override:
        scale = cy_override / native_cy
        return int(native_cx * scale), cy_override

    # Fit within max bounds preserving aspect ratio
    scale_x = max_cx / native_cx if native_cx > max_cx else 1.0
    scale_y = max_cy / native_cy if native_cy > max_cy else 1.0
    scale = min(scale_x, scale_y)
    return int(native_cx * scale), int(native_cy * scale)


def generate_pic_xml(rid, name, x, y, cx, cy, round_radius=0):
    """Generate <p:pic> XML element for a slide."""
    # Geometry: rounded rect or plain rect
    if round_radius > 0:
        geom = (
            f'<a:prstGeom prst="roundRect">'
            f'<a:avLst><a:gd name="adj" fmla="val {round_radius}"/></a:avLst>'
            f'</a:prstGeom>'
        )
    else:
        geom = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'

    return f'''<!-- {name} -->
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="0" name="{name}"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rid}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    {geom}
  </p:spPr>
</p:pic>'''


def main():
    parser = argparse.ArgumentParser(description='Add image to unpacked PPTX slide')
    parser.add_argument('unpacked_dir', help='Path to unpacked PPTX directory')
    parser.add_argument('slide_num', type=int, help='Slide number (1-based)')
    parser.add_argument('image_path', help='Path to image file')
    parser.add_argument('--x', type=int, default=457200, help='X position in EMU')
    parser.add_argument('--y', type=int, default=1200000, help='Y position in EMU')
    parser.add_argument('--cx', type=int, default=None, help='Width in EMU (overrides auto)')
    parser.add_argument('--cy', type=int, default=None, help='Height in EMU (overrides auto)')
    parser.add_argument('--max-cx', type=int, default=8229600, help='Max width in EMU')
    parser.add_argument('--max-cy', type=int, default=3200000, help='Max height in EMU')
    parser.add_argument('--name', default='Image', help='Shape name')
    parser.add_argument('--round', type=int, default=0, dest='round_radius',
                        help='Corner radius for rounded crop (EMU)')
    args = parser.parse_args()

    unpacked = Path(args.unpacked_dir)
    image_path = Path(args.image_path)

    # Validate paths
    if not unpacked.exists():
        print(f"ERROR: Unpacked directory not found: {unpacked}", file=sys.stderr)
        sys.exit(1)
    if not image_path.exists():
        print(f"ERROR: Image file not found: {image_path}", file=sys.stderr)
        sys.exit(1)

    slide_xml = unpacked / 'ppt' / 'slides' / f'slide{args.slide_num}.xml'
    if not slide_xml.exists():
        print(f"ERROR: Slide not found: {slide_xml}", file=sys.stderr)
        sys.exit(1)

    ext = image_path.suffix.lower()
    if ext not in CONTENT_TYPES:
        print(f"ERROR: Unsupported image format: {ext}", file=sys.stderr)
        sys.exit(1)

    # Get image dimensions
    img_w, img_h = get_image_dimensions(str(image_path))
    if not img_w or not img_h:
        print(f"ERROR: Could not read image dimensions from {image_path}", file=sys.stderr)
        sys.exit(1)
    print(f"Image: {img_w}x{img_h}px ({ext})")

    # Calculate display dimensions
    cx, cy = calculate_dimensions(img_w, img_h, args.max_cx, args.max_cy, args.cx, args.cy)
    print(f"Display: {cx}x{cy} EMU ({cx/914400:.2f}\"x{cy/914400:.2f}\")")

    # Copy image to ppt/media/
    media_dir = unpacked / 'ppt' / 'media'
    media_dir.mkdir(parents=True, exist_ok=True)
    idx = find_next_media_index(media_dir)
    dest_name = f'image{idx}{ext}'
    dest_path = media_dir / dest_name
    shutil.copy2(image_path, dest_path)
    print(f"  Copied: {image_path.name} -> ppt/media/{dest_name}")

    # Update [Content_Types].xml
    ct_path = unpacked / '[Content_Types].xml'
    if ct_path.exists():
        add_content_type(ct_path, ext)

    # Add relationship to slide .rels
    rels_dir = unpacked / 'ppt' / 'slides' / '_rels'
    rels_path = rels_dir / f'slide{args.slide_num}.xml.rels'
    rid, _ = find_next_rid(rels_path)
    add_relationship(rels_path, rid, f'../media/{dest_name}')

    # Generate and print XML snippet
    pic_xml = generate_pic_xml(rid, args.name, args.x, args.y, cx, cy, args.round_radius)
    print(f"\n=== Paste this into slide{args.slide_num}.xml (inside <p:spTree>) ===\n")
    print(pic_xml)
    print(f"\n=== Done. Image ready as {rid} in slide {args.slide_num} ===")

    return pic_xml


if __name__ == '__main__':
    main()
