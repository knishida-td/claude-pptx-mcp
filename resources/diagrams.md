# ネイティブ図解パターン

pptx_generate で作成したスライドにpython-pptxでネイティブオブジェクト（矩形・円・矢印等）を追加する際の設計パターン。

ヘルパー関数: `scripts/diagram_helpers.py`

## DiagramKit クラス

```python
from diagram_helpers import DiagramKit

kit = DiagramKit('/tmp/base.pptx')
sl = kit.slide(5)  # スライド5（1始まり）

# テンプレシェイプ以外を削除して図解を追加
kit.clear_custom_shapes(sl)
# or bigtextの空プレースホルダーのみ削除
kit.clear_bigtext_placeholders(sl)
```

## フォントサイズ階層（図解スライド）

| 役割 | サイズ | 用途例 |
|------|--------|--------|
| ヒーロー | 18pt bold | ステップラベル（入力/変換/出力）、セクション見出し |
| 見出し | 13pt bold | カードタイトル、項目名、質問 |
| 本文 | 11pt | 説明文、メッセージ、サブラベル |
| キャプション | 9pt | 注釈、小ラベル、フェーズ名 |

pptx_generate 生成スライド（本文14pt）との差は意図的。図解スライドは情報密度が高いため小さめ。

## 図解パターン一覧

### 1. 横フロー（process-flow）
3ステップまで。ステップ間に矢印を配置。

```python
from diagram_helpers import pattern_horizontal_flow

steps = [
    ('STEP 1', '観察する', kit.ORANGE, kit.WHITE),
    ('STEP 2', '分解する', kit.YELLOW, kit.DARK),
    ('STEP 3', '表現する', kit.TEAL, kit.WHITE),
]
pattern_horizontal_flow(kit, sl, y0=1.5, steps=steps)
```

### 2. カードグリッド（card-grid）
2×3、1×4 等の均等カード配置。

```python
from diagram_helpers import pattern_card_grid

cards = [
    ('結論だけ型', '理由を聞かれると詰まる'),
    ('補足過剰型', '説明が止まらず肥大化'),
    ('抽象逃げ型', '「いい感じ」で濁す'),
]
pattern_card_grid(kit, sl, y0=1.2, cards=cards, cols=3, card_h=1.5)
```

### 3. 2カラム比較（two-column）
左右に枠線付きボックス + 丸アイコン + 箇条書き。

```python
from diagram_helpers import pattern_two_column

pattern_two_column(kit, sl, y0=1.3,
    left_title='個人', left_items=[
        ('評価されない', 'アイデアが伝わらない'),
        ('信頼を失う', '曖昧な報告で不信感'),
    ],
    right_title='組織', right_items=[
        ('手戻り多発', '認識ズレで作業やり直し'),
        ('属人化', '暗黙知が共有されない'),
    ],
    left_color=kit.ORANGE, right_color=kit.YELLOW)
```

### 4. チップ一覧（chip-grid）
カテゴリラベル + N個の小さなチップ。30語一覧など。

```python
from diagram_helpers import pattern_chip_grid

categories = [
    ('感嘆', kit.YELLOW, kit.DARK, ['見事','圧巻','驚異的','壮大','素晴らしい','お見事']),
    ('感情', kit.RED_LIGHT, kit.WHITE, ['鳥肌が立つ','息を呑む','衝撃的','胸が震える','感動的','心を打つ']),
]
pattern_chip_grid(kit, sl, y0=1.5, categories=categories, chips_per_row=6)
```

### 5. 人物カード（person-cards）
ファクトバー + N人分の人物カード。

```python
from diagram_helpers import pattern_person_cards

persons = [
    ('社長', '方向性は？', '既存依存に偏り\n新規立て直し急務', kit.ORANGE, kit.WHITE),
    ('営業', '何をする？', '新規アポ\n週5件に引上げ', kit.YELLOW, kit.DARK),
    ('経理', '内訳は？', '既存+30%\n新規-15%', kit.TEAL, kit.WHITE),
]
pattern_person_cards(kit, sl, y0=1.2, fact_text='売上 前年比120%', persons=persons)
```

## バリデーション

```python
# 不要改行チェック（全スライド）
issues = kit.check_all_overflow()
for iss in issues:
    print(f"S{iss['slide']} {iss['shape']}: +{iss['overflow_pct']:.0f}% | {iss['text']}")

# 境界外チェック（全スライド）
bounds = kit.check_bounds()
for b in bounds:
    print(f"S{b['slide']} {b['shape']}: right={b['right']:.1f} bottom={b['bottom']:.1f}")

# KeyMsg背景の確認・追加
for sl in kit.prs.slides:
    kit.ensure_keymsg_bg(sl)
```

## 縦中央配置の計算

```python
# コンテンツ高さから理想的なy座標を算出
content_h = 3.0  # inch
y0 = kit.ideal_top(content_h)  # → 1.365 inch
```

## よくあるバグ

| バグ | 原因 | 対策 |
|------|------|------|
| シェイプ幅が640,080インチ | `Inches()`の二重適用 | 座標は生のfloat（inch）で管理し、関数内で1回だけ`Inches()`変換 |
| KeyMsgが背景なしに | クリーンアップでShape 6を削除 | `TEMPLATE_SHAPES`に含まれるシェイプは削除禁止。`ensure_keymsg_bg()`で事後チェック |
| 左右バランスが悪い | 数学的中央に配置 | 暗い・大きい要素は視覚的に重い。重い側のマージンを広くとる |
| フォントサイズがバラバラ | スライドごとにアドホックに決定 | 上記の4段階階層（18/13/11/9pt）を厳守 |
