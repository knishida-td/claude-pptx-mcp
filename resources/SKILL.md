---
name: ppt
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file (even if the extracted content will be used elsewhere, like in an email or summary); editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename, regardless of what they plan to do with the content afterward. If a .pptx file needs to be opened, created, or touched, use this skill."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX Skill

## 絶対ルール（全作成方式共通）

### 画像は必ず入れる

テキストとチャートだけのプレゼンは禁止。意味のあるスライドには必ず画像を差し込む。
- 企業概要 → 実際の商品パッケージ写真
- チャネル提案 → 販路の現場写真（コンビニ棚、ジム、店舗等）
- サービス提案 → サービス利用シーンの写真
- 画像はレイアウト設計の段階から配置を決める。後付けで余白に入れるのはNG

### 実物写真を使う

対象企業・ブランドの実際の商品写真をウェブ検索（楽天・Amazon等）で取得して使う。
ストック写真（汎用的な人物・風景）で代替しない。
「この画像を削除したらスライドの伝達力が落ちるか？」がYESになる画像のみ使用。

### 出力ファイルのバージョン管理

PPTXファイルを上書きしない。変更のたびに新バージョンで出力する。
```
❌ VALX事業改善提案_slidekit.pptx（同名上書き）
✅ VALX事業改善提案_slidekit_v1.pptx → _v2.pptx → _v3.pptx
```
生成スクリプトも同様にバージョンを上げる（`generate-xxx-v2.js` → `-v3.js`）。

---

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit or create from template | Read [editing.md](editing.md) |
| Create from scratch | Read [pptxgenjs.md](pptxgenjs.md) |
| Create with OOXML SlideKit | Read [slidekit.md](slidekit.md) |

---

## Reading Content

```bash
# Text extraction
python -m markitdown presentation.pptx

# Visual overview
python scripts/thumbnail.py presentation.pptx

# Raw XML
python scripts/office/unpack.py presentation.pptx unpacked/
```

---

## Editing Workflow

**Read [editing.md](editing.md) for full details.**

1. Analyze template with `thumbnail.py`
2. Unpack → manipulate slides → edit content → clean → pack

---

## Creating from Scratch

**Read [pptxgenjs.md](pptxgenjs.md) for full details.**

Use when no template or reference presentation is available.

### 提案資料は20枚以上

提案書・改善提案は **最低20枚以上** で厚みを持たせる。
構成テンプレートは [japanese-market-rules.md](japanese-market-rules.md) の「提案資料の標準構成」を参照。
実績のある参照スクリプト: `/tmp/bijin-ec-proposal-v5.js`（24枚・SlideKit準拠・実データ反映済み）

---

## Design Ideas

**⚠️ SlideKitを使用する場合、以下のDesign Ideasセクションは完全に無視すること。**
SlideKitには独自のデザインシステム（カラーパレット、フォント、レイアウトパターン）が定義されている。
SlideKit使用時にこのセクションのカラーパレット・フォント・ダーク背景等を適用すると、
デザインシステムが壊れる。SlideKit使用時は [slidekit.md](slidekit.md) のルールのみに従う。

**Don't create boring slides.** Plain bullets on a white background won't impress anyone. Consider ideas from this list for each slide.

### Before Starting

- **Pick a bold, content-informed color palette**: The palette should feel designed for THIS topic. If swapping your colors into a completely different presentation would still "work," you haven't made specific enough choices.
- **Dominance over equality**: One color should dominate (60-70% visual weight), with 1-2 supporting tones and one sharp accent. Never give all colors equal weight.
- **Dark/light contrast**: Dark backgrounds for title + conclusion slides, light for content ("sandwich" structure). Or commit to dark throughout for a premium feel.
- **Commit to a visual motif**: Pick ONE distinctive element and repeat it — rounded image frames, icons in colored circles, thick single-side borders. Carry it across every slide.

### Color Palettes

Choose colors that match your topic — don't default to generic blue. Use these palettes as inspiration:

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

### For Each Slide

**Every slide needs a visual element** — image, chart, icon, or shape. Text-only slides are forgettable.

**Layout options:**
- Two-column (text left, illustration on right)
- Icon + text rows (icon in colored circle, bold header, description below)
- 2x2 or 2x3 grid (image on one side, grid of content blocks on other)
- Half-bleed image (full left or right side) with content overlay

**Data display:**
- Large stat callouts (big numbers 60-72pt with small labels below)
- Comparison columns (before/after, pros/cons, side-by-side options)
- Timeline or process flow (numbered steps, arrows)

**Visual polish:**
- Icons in small colored circles next to section headers
- Italic accent text for key stats or taglines

### Typography

**Choose an interesting font pairing** — don't default to Arial. Pick a header font with personality and pair it with a clean body font.

| Header Font | Body Font |
|-------------|-----------|
| Georgia | Calibri |
| Arial Black | Arial |
| Calibri | Calibri Light |
| Cambria | Calibri |
| Trebuchet MS | Calibri |
| Impact | Arial |
| Palatino | Garamond |
| Consolas | Calibri |

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | 14-16pt |
| Captions | 10-12pt muted |

### Spacing

- 0.5" minimum margins
- 0.3-0.5" between content blocks
- Leave breathing room—don't fill every inch

### Avoid (Common Mistakes)

- **Don't repeat the same layout** — 連続するスライドで同じレイアウトパターンを使わない。特に「左:大きい数字 + 右:3カード」「3列カード」等のテンプレートを複数スライドに量産するのはAIっぽさの最大の原因。ヘルパー関数でカードを量産するアプローチ自体を避ける
- **Don't use uniform big-number callouts** — 全提案スライドに72pt数字を置くパターンは典型的なAI生成。数字を大きくするのは本当に重要な1〜2箇所だけに限定する
- **Don't center body text** — left-align paragraphs and lists; center only titles
- **Don't skimp on size contrast** — titles need 36pt+ to stand out from 14-16pt body
- **Don't default to blue** — pick colors that reflect the specific topic
- **Don't mix spacing randomly** — choose 0.3" or 0.5" gaps and use consistently
- **Don't style one slide and leave the rest plain** — commit fully or keep it simple throughout
- **Don't create text-only slides** — add images, icons, charts, or visual elements; avoid plain title + bullets
- **Show relationships with arrows** — 要素を並べるだけでなく、矢印で因果関係・フロー・プロセスを明示する。`line: { endArrowType: "triangle" }` で矢印を描画。カード間の因果、施策のステップ、ファネルの流れ等に使う
- **Don't forget text box padding** — when aligning lines or shapes with text edges, set `margin: 0` on the text box or offset the shape to account for padding
- **Don't use low-contrast elements** — icons AND text need strong contrast against the background; avoid light text on light backgrounds or dark text on dark backgrounds
- **Accent lines under titles are OK** — dual-color accent lines (e.g. red + navy) under headers are a clean design pattern, not an AI hallmark
- **NEVER use dark-mode backgrounds** — 黒・濃紺・ダークグレーなどの暗い背景色は禁止。ライト系の背景のみ使用すること
- **NEVER spam drop-shadow cards** — ドロップシャドウ付きカードの連打禁止。`shadow` プロパティは原則使わない
- **NEVER use left-edge accent bars on cards** — カード左端の細いカラーバー（`w: 0.06`）はAI生成スライドの典型パターン。使わない
- **NEVER use top-edge color lines on cards** — カード上端の細いカラーライン（`h: 0.06`）も同様。使わない
- **NEVER use English-label headings** — 見出しに英語ラベル（"Overview", "Key Takeaways" 等）を入れない。日本語資料では日本語の見出しのみ使用すること
- **Don't over-color-code** — カードごとに色を変えてタイトルを着色するパターンもAIっぽい。色は2〜3色に抑え、意味のある箇所だけに使う
- **Don't use em dashes（ — ）** — AIが多用する記号。句点で文を区切るか、括弧を使う。「A — B」→「A。B」「A（B）」
- **Don't write AI-like short declarative sentences** — 「攻めは動いた。回収が追いついていない」のような体言止め短文の連打はAIっぽい。「攻めは動いたが、回収が追いついていない点が課題」のように接続詞で繋げる方が自然
- **Don't use sub-10pt fonts for body text** — 本文に8ptは小さすぎる。本文は11pt以上、補足注釈は9pt以上を基準にする
- **NEVER use neon-color accents** — ネオンカラー（蛍光グリーン、蛍光ピンク等）のアクセント禁止。落ち着いたトーンのアクセントカラーを選ぶこと
- **NEVER put emojis in slides** — スライド内に絵文字を入れない。アイコンが必要な場合はシェイプや画像で表現すること
- **NEVER substitute icon+card grids for real diagrams** — アイコン＋カードの羅列で図解の代わりにしない。関係性やフローを示すなら矢印・コネクタ等を使った本物の図を作ること

---

## QA (Required)

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Approach QA as a bug hunt, not a confirmation step. If you found zero issues on first inspection, you weren't looking hard enough.

### Content QA

```bash
python -m markitdown output.pptx
```

Check for missing content, typos, wrong order.

**When using templates, check for leftover placeholder text:**

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

If grep returns results, fix them before declaring success.

### Visual QA

**⚠️ USE SUBAGENTS** — even for 2-3 slides. You've been staring at the code and will see what you expect, not what's there. Subagents have fresh eyes.

Convert slides to images (see [Converting to Images](#converting-to-images)), then use this prompt:

```
Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Low-contrast text (e.g., light gray text on cream-colored background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without a contrasting circle)
- Text boxes too narrow causing excessive wrapping
- Leftover placeholder content

For each slide, list issues or areas of concern, even if minor.

Read and analyze these images:
1. /path/to/slide-01.jpg (Expected: [brief description])
2. /path/to/slide-02.jpg (Expected: [brief description])

Report ALL issues found, including minor ones.
```

### Verification Loop

1. Generate slides → Convert to images → Inspect
2. **List issues found** (if none found, look again more critically)
3. Fix issues
4. **Re-verify affected slides** — one fix often creates another problem
5. Repeat until a full pass reveals no new issues

**Do not declare success until you've completed at least one fix-and-verify cycle.**

---

## Converting to Images

Convert presentations to individual slide images for visual inspection:

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

This creates `slide-01.jpg`, `slide-02.jpg`, etc.

To re-render specific slides after fixes:

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

---

## Dependencies

- `pip install "markitdown[pptx]"` - text extraction
- `pip install Pillow` - thumbnail grids
- `npm install -g pptxgenjs` - creating from scratch
- LibreOffice (`soffice`) - PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`) - PDF to images
