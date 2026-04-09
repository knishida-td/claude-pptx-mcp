#!/usr/bin/env python3
"""
画像検索・ダウンロード・サイズ計算ヘルパー

2つの画像取得方法をサポート:
1. ぱくたそタグ検索（日本人ビジネス写真）
2. 任意URLからのダウンロード（商品画像、WebSearchで見つけたURL等）

Usage:
    from image_search import search_pakutaso, download_image, fit_image_size

    # ぱくたそから検索してダウンロード
    path = search_and_download("リモートワーク", "/tmp/remote.jpg")

    # 任意URLからダウンロード（商品画像等）
    path = download_image("https://example.com/product.jpg", "/tmp/product.jpg")

    # アスペクト比維持でサイズ計算
    w, h = fit_image_size(path, max_w=2.7, max_h=1.5)

    # PPTXに挿入（widthのみ指定でアスペクト比維持）
    from pptx.util import Inches
    slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w))
"""

import json
import os
import re
import sys
import urllib.request
import urllib.parse
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    Image = None


def search_pakutaso(query: str, num_results: int = 5) -> list[str]:
    """ぱくたそのタグページから画像CDN URLを取得。

    queryの各単語をタグとして検索し、見つかったCDN URLを返す。
    複数単語の場合は最初にヒットしたタグの結果を使用。

    Returns:
        CDN URL (Sサイズ) のリスト
    """
    keywords = query.split()
    all_urls = []

    for keyword in keywords:
        tag_url = ('https://www.pakutaso.com/tag/'
                   + urllib.parse.quote(keyword) + '.html')
        req = urllib.request.Request(tag_url, headers={
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                           'AppleWebKit/537.36 (KHTML, like Gecko) '
                           'Chrome/120.0.0.0 Safari/537.36'
        })

        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                html = resp.read().decode('utf-8', errors='ignore')
        except Exception as e:
            print(f"Tag search failed for '{keyword}': {e}", file=sys.stderr)
            continue

        # CDN URLを抽出（重複除去、順序保持）
        pattern = r'https://user0514\.cdnw\.net/shared/img/thumb/([A-Za-z0-9_]+)_TP_V4\.jpg'
        names = list(dict.fromkeys(re.findall(pattern, html)))

        for name in names[:num_results]:
            cdn = f"https://user0514.cdnw.net/shared/img/thumb/{name}_TP_V4.jpg"
            if cdn not in all_urls:
                all_urls.append(cdn)

        if all_urls:
            break  # 最初にヒットしたタグで十分

    return all_urls[:num_results]


def get_pakutaso_cdn_url(page_url: str, size: str = 'S') -> str | None:
    """ぱくたそのページURLからCDN画像URLを取得。

    Args:
        page_url: ぱくたその写真ページURL
        size: 'S' (800px), 'M' (1600px), 'L' (原寸)
    """
    req = urllib.request.Request(page_url, headers={
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/120.0.0.0 Safari/537.36'
    })

    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"Failed to fetch {page_url}: {e}", file=sys.stderr)
        return None

    # CDN URLパターン: https://user0514.cdnw.net/shared/img/thumb/XXXXX.jpg
    # S: _TP_V4.jpg, M: _TP_V.jpg, L: .jpg (サフィックスなし)
    pattern = r'https://user0514\.cdnw\.net/shared/img/thumb/([A-Za-z0-9_]+)(?:_TP_V4|_TP_V)?\.jpg'
    match = re.search(pattern, html)
    if not match:
        return None

    base_name = match.group(1)
    base_url = f"https://user0514.cdnw.net/shared/img/thumb/{base_name}"

    if size == 'S':
        return f"{base_url}_TP_V4.jpg"
    elif size == 'M':
        return f"{base_url}_TP_V.jpg"
    else:  # L
        return f"{base_url}.jpg"


def download_image(cdn_url: str, save_path: str) -> str | None:
    """CDN URLから画像をダウンロード。"""
    req = urllib.request.Request(cdn_url, headers={
        'User-Agent': 'Mozilla/5.0'
    })

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read()
            with open(save_path, 'wb') as f:
                f.write(data)
        return save_path
    except Exception as e:
        print(f"Download failed: {e}", file=sys.stderr)
        return None


def get_image_size(path: str) -> tuple[int, int]:
    """画像の幅と高さを取得。PILがなければバイナリ解析。"""
    if Image:
        img = Image.open(path)
        return img.size

    # PILなしのフォールバック: JPEGヘッダー解析
    with open(path, 'rb') as f:
        data = f.read()

    # JPEG SOF0/SOF2マーカーから取得
    i = 0
    while i < len(data) - 1:
        if data[i] == 0xFF:
            marker = data[i + 1]
            if marker in (0xC0, 0xC2):  # SOF0, SOF2
                h = (data[i + 5] << 8) | data[i + 6]
                w = (data[i + 7] << 8) | data[i + 8]
                return (w, h)
            elif marker == 0xD8 or marker == 0xD9:
                i += 2
            else:
                length = (data[i + 2] << 8) | data[i + 3]
                i += 2 + length
        else:
            i += 1

    raise ValueError(f"Cannot determine image size: {path}")


def fit_image_size(path: str, max_w: float = None, max_h: float = None
                   ) -> tuple[float, float]:
    """アスペクト比を維持してmax_w/max_h内に収まるサイズを計算。

    Args:
        path: 画像ファイルパス
        max_w: 最大幅 (inch)
        max_h: 最大高さ (inch)

    Returns:
        (width_inch, height_inch) アスペクト比維持済み
    """
    pw, ph = get_image_size(path)
    ratio = pw / ph

    if max_w and max_h:
        # 両方指定: ボックス内にフィット
        w = max_w
        h = w / ratio
        if h > max_h:
            h = max_h
            w = h * ratio
    elif max_w:
        w = max_w
        h = w / ratio
    elif max_h:
        h = max_h
        w = h * ratio
    else:
        # デフォルト: 幅3インチ
        w = 3.0
        h = w / ratio

    return (w, h)


def search_and_download(query: str, save_path: str, size: str = 'S',
                        index: int = 0) -> str | None:
    """検索→ダウンロードを一括実行。

    Args:
        query: 検索キーワード（例: "リモートワーク"）
        save_path: 保存先パス
        size: 'S' (800px), 'M' (1600px), 'L' (原寸)
        index: 検索結果の何番目を使うか（0始まり）

    Returns:
        保存先パス or None
    """
    cdn_urls = search_pakutaso(query, num_results=index + 3)
    if not cdn_urls:
        print(f"No results for: {query}", file=sys.stderr)
        return None

    # サイズに応じてURL変換
    if index < len(cdn_urls):
        url = cdn_urls[index]
    else:
        url = cdn_urls[0]

    if size == 'M':
        url = url.replace('_TP_V4.jpg', '_TP_V.jpg')
    elif size == 'L':
        url = url.replace('_TP_V4.jpg', '.jpg')

    path = download_image(url, save_path)
    if path:
        w, h = get_image_size(path)
        print(f"Downloaded: {query} -> {path} ({w}x{h})")
        return path

    print(f"Failed to download for: {query}", file=sys.stderr)
    return None


# === CLI ===
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='ぱくたそ画像検索・ダウンロード')
    parser.add_argument('query', help='検索キーワード')
    parser.add_argument('-o', '--output', default='/tmp/pakutaso_image.jpg',
                        help='保存先パス')
    parser.add_argument('-s', '--size', choices=['S', 'M', 'L'], default='S',
                        help='画像サイズ (S=800px, M=1600px, L=原寸)')
    parser.add_argument('--max-w', type=float, help='最大幅 (inch)')
    parser.add_argument('--max-h', type=float, help='最大高さ (inch)')
    args = parser.parse_args()

    path = search_and_download(args.query, args.output, args.size)
    if path and (args.max_w or args.max_h):
        w, h = fit_image_size(path, args.max_w, args.max_h)
        print(f"Recommended size: {w:.2f} x {h:.2f} inch")
