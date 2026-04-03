# SlideKit: OOXML直接生成によるスライド作成

pptxgenjsやpython-pptxを使わず、OOXML XMLを直接制御してPPTXを生成する。

単位はEMU (English Metric Units)。1インチ=914400, 1pt=12700。16:9スライド=9144000×5143500。

---

## デザインシステム

### カラーパレット

| 役割 | HEX | 用途 |
|---|---|---|
| 背景 | `F5F5F5` | 全スライド統一 |
| タイトル | `222222` | スライドタイトル・大見出し |
| 本文 | `333333` | 箇条書き・説明テキスト |
| 補足 | `666666` | グリッドの副テキスト等 |
| ミュート | `AAAAAA` | ページ番号・マーカー・セカンダリラベル |
| プライマリ | `EF4823` | 見出し・KeyMsg・アクセントバー・矢印 |
| セカンダリ | `FCBF17` | YellowLine専用 |
| KeyMsg背景 | `FFF5F0` | KeyMsgBarの背景 |
| セパレーター | `EEEEEE` | 細い区切り線 |
| デバイダー | `DDDDDD` | 太い縦区切り線 |

### フォント・サイズ

全テキスト統一: `Hiragino Kaku Gothic Pro W3`

| 要素 | sz | bold | 色 |
|---|---|---|---|
| スライドタイトル | 2200 | b="1" | 222222 |
| ヒーローテキスト | 2800-4000 | b="1" | EF4823 |
| セクション見出し | 1600 | b="1" | EF4823 |
| 本文 | 1400 | — | 333333 |
| キーメッセージ | 1800 | b="1" | EF4823 |
| ページ番号 | 900 | — | AAAAAA |
| 副テキスト | 1200 | — | 666666 |

---

## 共通コンポーネント

全コンテンツスライドで固定位置。XMLテンプレートはコピーして使う。

| 要素 | x | y | cx | cy | 備考 |
|---|---|---|---|---|---|
| Title | 457200 | 354510 | 8229600 | 411120 | sz=2200 bold, anchor=ctr |
| RedLine | 457200 | 784350 | 3886200 | 31680 | fill=EF4823 |
| YellowLine | 4343400 | 784350 | 4343400 | 31680 | fill=FCBF17 |
| KeyMsgBg | 457200 | 4423590 | 8229240 | 365400 | roundRect, fill=FFF5F0 |
| KeyMsgText | 457200 | 4423590 | 8229240 | 365400 | sz=1800 bold, EF4823, 中央揃え |
| PageNum | 8412480 | 4880790 | 456840 | 228240 | sz=900, AAAAAA, 右揃え |

**ヘッダー下端**: y=816030 (RedLine/YellowLine bottom)
**KeyMsg上端**: y=4423590
**本体コンテンツ領域**: 816030〜4423590 = **3607560 EMU**

---

## レイアウトパターン

### A: タイトル / B: セクション扉
KeyMsgなし。スライド全体(5143500)で縦中央配置。

Type A/Bも Type D と同様、コンテンツ高さを実測して `bTop = (SH - contentH) / 2` で算出する。
ハードコードした固定値（例: `bTop = 1.8`）は使わない。

### C: コンテンツスライド群
共通ヘッダー + KeyMsgフッター + 本体コンテンツ:

| パターン | 用途 | 主要要素 |
|---|---|---|
| **大見出し+補足** | インパクト主張 | BigText + SepLine + SubText |
| **左右2カラム** | 比較・対比 | Header×2 + Items×2 + Divider |
| **3カラムグリッド** | 並列3項目 | Subtitle + ColTitle×3 + ColBody×3 + Div×2 |
| **番号付きリスト** | 手順・要点 | Num + Item + Sep (×N) |
| **定義ブロック** | 定義・解説 | 赤バー + Heading + Body (×N) + Separator |
| **Before→After** | 変化表現 | LeftAccent + Before + Arrow + RightAccent + After |
| **2×2グリッド** | マトリクス | Grid×4 + HSep + VSep + CenterElement |
| **因果フロー** | 原因分析 | Problem + Arrow + Factors (×N) |
| **プロセスフロー** | 工程表示（3ステップまで） | Step(roundRect)×N + Arrow×(N-1) + Details |
| **番号付き縦リスト** | 4ステップ以上の工程 | Num(circle) + Title + Desc (×N) + Sep |
| **A/B選択肢** | 選択提示 | Label + Subtitle + Body (×2) + Divider |

### 矢印の方向

**矢印の方向はコンテンツの流れと一致させる。例外なし。**

| レイアウト | 矢印の方向 |
|---|---|
| Before→After が左右に並ぶ | 横矢印（→） |
| Before→After が上下に並ぶ | 縦矢印（↓） |
| プロセスフローが横並び | 横矢印（→） |
| プロセスフローが縦並び | 縦矢印（↓） |

横に並んでいるのに縦矢印、縦に並んでいるのに横矢印は「気持ち悪い」ので絶対にやらない。

### プロセスフローの注意

**横並びプロセスフロー（circle + title + desc）は3ステップが上限。**
4ステップ以上は横幅が足りず、descテキストが詰まって崩れる。
4ステップ以上 → **番号付き縦リスト**パターンに切り替える。

### D: CTA / エンド
KeyMsgなし。**縦中央 AND 左右中央に配置。固定値のx決め打ち禁止。**

Type Dスライドはヘッダー・KeyMsgがないため、スライド全体（10"×5.625"）に対して中央配置する。

**手順（省略禁止）:**

```javascript
const SW = 10;   // スライド幅
const SH = 5.625; // スライド高さ

// ── STEP 1: 縦中央 ──
// 全要素のh + gapを積み上げてcontentHを算出
const contentH = titleH + gapToLine + lineH + gapToSteps + (steps * stepH) + gapToFooter + footerH;
const baseY = (SH - contentH) / 2;

// ── STEP 2: 左右中央 ──
// 1行の中で最も左の要素〜最も右の要素の右端までの総幅を算出
const blockW = circleW + gapToText + textW + gapToSub + subW;
const baseX = (SW - blockW) / 2;
// 各要素のxはbaseXからの相対位置で決める
// circle.x = baseX
// text.x   = baseX + circleW + gapToText
// sub.x    = baseX + circleW + gapToText + textW + gapToSub

// ── STEP 3: テキスト幅の検証 ──
// 最長テキストが1行に収まるかを確認（全角1文字≒0.22" @16pt）
// 収まらない場合はtextWを広げるかフォントを下げる。折り返し放置は禁止。
```

**禁止パターン:**
- `x: 1.8` 等の固定値でcircleやテキストの位置を決める → 左右中央にならない
- contentHを計算せずbaseYを `1.0` 等でハードコード → 上寄りになる
- テキストボックス幅を狭く設定し、閉じ括弧「）」等が次の行に折り返す → 必ず最長文字列で検証する

---

## レイアウトルール

### 縦中央配置

**全スライド必須。例外なし。上下の余白が均等でなければならない。**

縦中央 = 「ヘッダー下端〜コンテンツ上端」と「コンテンツ下端〜KeyMsg上端」の距離が等しいこと。
目視で「上に詰まっている」「下に寄っている」と感じたら中央になっていない。

```
# コンテンツスライド (タイプC)
tight_body_height = sum(各要素の適切なcy + 要素間のgap)
ideal_top = 816030 + (3607560 - tight_body_height) / 2
# 検証: (first_body_y - 816030) == (4423590 - last_body_bottom)

# セクション扉 (タイプA, B)
ideal_top = (5143500 - content_height) / 2
```

### テキストボックスのサイズ

`cy`は実際のテキスト量に合わせる。`anchor="t"` かつ `cy > 実テキスト高の2倍` → 縮小必須。

```
cy = (行数 × 行高) + ((行数-1) × 段落間スペース) + マージン(~50000)
```

**行高**: sz=1200→~218K, 1400→~254K, 1600→~291K, 2200→~400K, 2800→~508K, 4000→~728K

**段落間スペース**: `spcAft val="800"` = ~101600 EMU

### 要素間マージン

| 要素間 | 最小gap |
|---|---|
| 同一セクション内（Heading→Body） | 80000 |
| セクション間（Sep前後それぞれ） | 100000 |
| リスト項目間（Sep前後それぞれ） | 100000 |
| 密接な要素（Label→Subtitle→Body） | 60000 |

**gap=0 は禁止**。

### コンテンツ充填率

本体コンテンツが利用可能領域に占める割合。**60-75%が理想**。50%未満ならcy・gapを拡大。

### テキスト折り返し制御

文字数がテキストボックス幅の95%以上 → `<a:br/>`で明示改行し、cyも行数分拡大。

**全角文字数/行**: cx=7315200でsz=1400→~41, sz=2400→~24, sz=2800→~20。cx=8229600でsz=1400→~46, sz=2400→~27, sz=2800→~23。

### セパレーターの重なり回避

コンテンツ要素と交差する場合、線を分割（最小gap=50000）:
```
HSepL: cx = icon_x - 50000 - left  /  HSepR: x = icon_x + icon_cx + 50000
VSepTop: cy = arrow_y - 50000 - top  /  VSepBot: y = arrow_y + arrow_cy + 50000
```

---

## 画像

### 画像の入れ方（最重要）

**画像は積極的に使いたい。ただし「画像ありき」でレイアウトしたスライドに入れる。**
テキストやデータ中心のスライドの余白に後付けで小さく入れるとダサくなる。

| 判断 | 理由 |
|------|------|
| ✅ レイアウト段階から画像スペースを確保したスライド | 商品紹介（右半分を画像エリアに）、顧客ジャーニー（右に商品写真）、ブランド紹介 |
| ❌ テキスト/データ中心レイアウトの余白に後付け | テーブル横の小さい商品画像、チャート隅の画像、ロードマップ脇の画像 |

**ルール**: 画像を入れたいなら、レイアウト設計の時点でそのスライドを「画像あり」で組む。
後から余白に入れるのはNG。

### 埋め込みスクリプト

```bash
python scripts/add_image.py UNPACKED_DIR SLIDE_NUM IMAGE_PATH [OPTIONS]
# --x/--y: 位置(EMU)  --cx/--cy: サイズ(EMU, 自動計算可)
# --max-cx/--max-cy: 最大制約  --name: シェイプ名  --round: 角丸半径
```

スクリプトが `ppt/media/` へのコピー、`[Content_Types].xml` 更新、`.rels` リレーションシップ追加、`<p:pic>` XMLスニペット出力を全自動で行う。

### ルール

- **アスペクト比を絶対に変えない**: `noChangeAspect="1"` 必須
- **解像度は可能な限り高いものを使う**: 150DPI以上
- **テキストとの重なり防止**: 画像を入れたら他の要素を画像と重ならないように調整
- **縦中央配置に含める**: 画像もボディコンテンツの一部として計算

---

## ファイル構造と破損防止

### 必須ファイル

```
[Content_Types].xml / _rels/.rels
ppt/presentation.xml / ppt/_rels/presentation.xml.rels / ppt/presProps.xml
ppt/theme/theme1.xml
ppt/slideMasters/slideMaster1.xml (+_rels)
ppt/slideLayouts/slideLayout1.xml (+_rels)
ppt/slides/slide{N}.xml (+_rels) / docProps/app.xml / docProps/core.xml
```

### 破損防止チェックリスト

- Content_Typesに**同一Extensionを重複追加しない**
- presentation.xmlの**スライドサイズを変更しない**
- 画像の追加/削除時は**media・.rels・Content_Typesの3箇所を必ず整合**させる
- リパック前に**全XMLをET.parse()で検証**
- 画像0枚なら**空のmedia/ディレクトリを残さない**

---

## 作成ワークフロー

1. **構成決定** → レイアウトパターン選択
2. **XML生成** → 共通コンポーネントテーブルから座標をコピー
3. **レイアウト調整** → 縦中央配置・充填率・マージン・折り返し確認
4. **パッケージング** → `cd unpacked/ && zip -r -X ../output.pptx . -x ".*" "__MACOSX/*"`
5. **QA** → [SKILL.md](SKILL.md) のQAセクション参照
