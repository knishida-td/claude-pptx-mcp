#!/bin/bash
# PostToolUse hook: PPTX生成後に自動でサムネイル生成してQAリマインダーを表示
# Bashコマンドの出力に .pptx が含まれていたら発火する

# stdin からツール結果を読む
INPUT=$(cat)

# .pptx ファイルパスを検出（node xxx.js の出力 "Done: /path/to.pptx" など）
PPTX_PATH=$(echo "$INPUT" | grep -oE '/[^ ]*\.pptx' | head -1)

if [ -z "$PPTX_PATH" ]; then
  echo "$INPUT"
  exit 0
fi

# PPTXファイルが実在するか確認
if [ ! -f "$PPTX_PATH" ]; then
  echo "$INPUT"
  exit 0
fi

# QAリマインダーを追加して出力
cat <<EOF
$INPUT

───── PPTX QAリマインダー ─────
生成ファイル: $PPTX_PATH
次のステップ:
1. soffice → pdftoppm で画像化
2. サブエージェントで全スライド目視QA
3. KeyMsg 28文字以内を確認
4. 問題ゼロまで修正→再確認ループ
──────────────────────────────
EOF
