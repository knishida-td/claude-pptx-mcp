#!/bin/bash
# Slidekit PPTX生成スクリプトのルール違反を自動検出
# Usage: validate-slidekit.sh <path-to-js-file>

set -euo pipefail

FILE="${1:?Usage: validate-slidekit.sh <js-file>}"
ERRORS=0

err() { echo "  ❌ $1"; ((ERRORS++)); }
ok()  { echo "  ✅ $1"; }

echo "=== Slidekit バリデーション: $(basename "$FILE") ==="
echo ""

# 1. カラーパレット: EF4823/FCBF17が定義されているか
echo "[1] カラーパレット"
if grep -q 'EF4823' "$FILE" && grep -q 'FCBF17' "$FILE"; then
  ok "EF4823 (primary) / FCBF17 (secondary) を使用"
else
  err "EF4823/FCBF17 が見つからない。Slidekitカラーパレットを使用していない"
fi

# 2. ヘルパー関数
echo "[2] ヘルパー関数"
for fn in addHeader addKeyMsg addPageNum centerY; do
  if grep -q "$fn" "$FILE"; then
    ok "$fn を使用"
  else
    err "$fn が見つからない"
  fi
done

# 3. sizing: cover の禁止
echo "[3] sizing: cover"
COVER_COUNT=$(grep -c 'sizing.*cover' "$FILE" 2>/dev/null || true)
if [ "$COVER_COUNT" -gt 0 ]; then
  err "sizing: cover が ${COVER_COUNT} 箇所見つかった（アスペクト比が崩れる）"
  grep -n 'sizing.*cover' "$FILE" | head -5 | while read -r line; do echo "      $line"; done
else
  ok "sizing: cover なし"
fi

# 4. アスペクト比計算: fitImage または手動計算が存在するか
echo "[4] アスペクト比計算"
if grep -q 'fitImage\|aspectRatio\|ratio\|origWidth.*origHeight' "$FILE"; then
  ok "アスペクト比計算コードあり"
else
  if grep -q 'addImage' "$FILE"; then
    err "画像を使用しているがアスペクト比計算コードが見つからない"
  else
    ok "画像未使用（スキップ）"
  fi
fi

# 5. #付きカラーコード（PptxGenJSでは禁止）
echo "[5] PptxGenJS カラーコード"
# color: "#FFFFFF" 等のパターンを検出。ただしreact-iconsのSVGレンダリング用は除外
HEX_HASH=$(grep -E '(color|fill).*"#[0-9A-Fa-f]' "$FILE" 2>/dev/null | grep -cv 'renderIconSvg\|iconToBase64\|React.createElement' || true)
if [ "$HEX_HASH" -gt 0 ]; then
  err "#付きカラーコードが ${HEX_HASH} 箇所（PptxGenJSではファイル破損の原因）"
  grep -nE '(color|fill).*"#[0-9A-Fa-f]' "$FILE" | head -5 | while read -r line; do echo "      $line"; done
else
  ok "#付きカラーコードなし"
fi

# 6. スライド枚数（addSlideの数をカウント）
echo "[6] スライド枚数"
SLIDE_COUNT=$(grep -c 'addSlide()' "$FILE" 2>/dev/null || true)
if [ "$SLIDE_COUNT" -ge 20 ]; then
  ok "${SLIDE_COUNT} スライド"
elif [ "$SLIDE_COUNT" -ge 10 ]; then
  err "${SLIDE_COUNT} スライド（20枚以上推奨）"
else
  err "${SLIDE_COUNT} スライド（少なすぎる）"
fi

# 7. オプションオブジェクトの再利用（shadow等）
echo "[7] オプション再利用"
# 同じ変数名のshadowを複数回渡していないかチェック
REUSED=$(grep -cE 'shadow:\s*shadow[,\s}]' "$FILE" 2>/dev/null || true)
if [ "$REUSED" -gt 1 ]; then
  err "shadowオブジェクトが再利用されている可能性（毎回新規作成が必要）"
else
  ok "オプション再利用なし"
fi

# 8. 画像のloadImage/loadImageOrPlaceholderが存在するか
echo "[8] 画像ヘルパー"
if grep -q 'loadImage\|loadImageOrPlaceholder' "$FILE"; then
  ok "画像読み込みヘルパーあり"
else
  if grep -q 'addImage' "$FILE"; then
    err "addImageを使用しているが画像ヘルパー関数が未定義"
  else
    ok "画像未使用（スキップ）"
  fi
fi

# 結果サマリー
echo ""
echo "=== 結果 ==="
if [ "$ERRORS" -eq 0 ]; then
  echo "✅ 全チェック通過"
else
  echo "❌ ${ERRORS} 件のルール違反を検出"
fi
echo ""
exit "$ERRORS"
