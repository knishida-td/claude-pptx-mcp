#!/usr/bin/env bash
# claude-pptx-mcp ワンライナーインストール
# 使い方: curl -fsSL https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main/install.sh | bash

set -euo pipefail

SETTINGS_FILE="$HOME/.claude/settings.json"
MCP_ENTRY='"pptx":{"command":"npx","args":["-y","github:knishida-td/claude-pptx-mcp"]}'

echo "🔧 claude-pptx-mcp をインストールします..."

# settings.json がなければ作成
if [ ! -f "$SETTINGS_FILE" ]; then
  mkdir -p "$(dirname "$SETTINGS_FILE")"
  echo '{}' > "$SETTINGS_FILE"
  echo "  settings.json を新規作成しました"
fi

# 既に設定済みならスキップ
if grep -q "claude-pptx-mcp" "$SETTINGS_FILE" 2>/dev/null; then
  echo "✅ 既にインストール済みです。Claude Codeを再起動してください。"
  exit 0
fi

# Node.js チェック
if ! command -v node &>/dev/null; then
  echo "❌ Node.js が見つかりません。先にインストールしてください:"
  echo "   https://nodejs.org/"
  exit 1
fi

# jq があれば使う、なければ Python で編集
if command -v jq &>/dev/null; then
  tmp=$(mktemp)
  jq --argjson pptx "{\"command\":\"npx\",\"args\":[\"-y\",\"github:knishida-td/claude-pptx-mcp\"]}" \
    '.mcpServers = (.mcpServers // {}) | .mcpServers.pptx = $pptx' \
    "$SETTINGS_FILE" > "$tmp" && mv "$tmp" "$SETTINGS_FILE"
else
  python3 -c "
import json, sys
path = '$SETTINGS_FILE'
with open(path) as f:
    data = json.load(f)
data.setdefault('mcpServers', {})
data['mcpServers']['pptx'] = {
    'command': 'npx',
    'args': ['-y', 'github:knishida-td/claude-pptx-mcp']
}
with open(path, 'w') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
    f.write('\n')
"
fi

# CLAUDE.md にPPTX生成ルールを追記
CLAUDE_MD="$HOME/.claude/CLAUDE.md"
PPTX_RULE='「資料作って」等のスライド作成依頼では、必ずPptxGenJS（Node.js）でPPTXファイルを生成すること。HTML出力禁止。pptx MCPサーバーのツールを使うこと。'

if [ ! -f "$CLAUDE_MD" ]; then
  echo "$PPTX_RULE" > "$CLAUDE_MD"
  echo "  CLAUDE.md を新規作成しました"
elif ! grep -q "PptxGenJS" "$CLAUDE_MD" 2>/dev/null; then
  printf '\n%s\n' "$PPTX_RULE" >> "$CLAUDE_MD"
  echo "  CLAUDE.md にPPTX生成ルールを追記しました"
else
  echo "  CLAUDE.md にPPTX生成ルールは設定済みです"
fi

# validate-slidekit.sh をインストール
VALIDATE_SRC=""
for candidate in \
  "$(cd "$(dirname "$0")" 2>/dev/null && pwd)/scripts/validate-slidekit.sh" \
  "$(npm root -g 2>/dev/null)/claude-pptx-mcp/scripts/validate-slidekit.sh"; do
  if [ -f "$candidate" 2>/dev/null ]; then
    VALIDATE_SRC="$candidate"
    break
  fi
done

if [ -n "$VALIDATE_SRC" ] && [ -f "$VALIDATE_SRC" ]; then
  mkdir -p "$HOME/.claude/scripts"
  cp "$VALIDATE_SRC" "$HOME/.claude/scripts/validate-slidekit.sh"
  chmod +x "$HOME/.claude/scripts/validate-slidekit.sh"
  echo "  validate-slidekit.sh をインストールしました"
fi

# PostToolUse hook を登録（PPTX生成後のQAリマインダー）
HOOKS_FILE="$HOME/.claude/hooks/hooks.json"
HOOK_SCRIPT_DIR="$HOME/.claude/scripts/hooks"
HOOK_SCRIPT="post-bash-pptx-qa.sh"

# install.sh が curl | bash で実行された場合、HOOK_SRC は取得できないので
# npx 経由でインストールされたパスから探す
HOOK_SRC=""
for candidate in \
  "$(cd "$(dirname "$0")" 2>/dev/null && pwd)/scripts/$HOOK_SCRIPT" \
  "$(npm root -g 2>/dev/null)/claude-pptx-mcp/scripts/$HOOK_SCRIPT"; do
  if [ -f "$candidate" 2>/dev/null ]; then
    HOOK_SRC="$candidate"
    break
  fi
done

# hookスクリプトをコピー
mkdir -p "$HOOK_SCRIPT_DIR"
if [ -n "$HOOK_SRC" ] && [ -f "$HOOK_SRC" ]; then
  cp "$HOOK_SRC" "$HOOK_SCRIPT_DIR/$HOOK_SCRIPT"
  chmod +x "$HOOK_SCRIPT_DIR/$HOOK_SCRIPT"
  echo "  QAリマインダーhookをインストールしました"
fi

# hooks.json に登録（既に登録済みならスキップ）
if [ -f "$HOOKS_FILE" ] && grep -q "post-bash-pptx-qa" "$HOOKS_FILE" 2>/dev/null; then
  echo "  hooks.json にQA hookは設定済みです"
else
  python3 -c "
import json, os
path = '$HOOKS_FILE'
if os.path.exists(path):
    with open(path) as f:
        data = json.load(f)
else:
    data = {'hooks': {}}

post = data.setdefault('hooks', {}).setdefault('PostToolUse', [])
if not any('post-bash-pptx-qa' in json.dumps(h) for h in post):
    post.append({
        'matcher': 'Bash',
        'hooks': [{'type': 'command', 'command': 'bash ~/.claude/scripts/hooks/post-bash-pptx-qa.sh'}],
        'description': 'PPTX生成検出時にQAリマインダーを表示'
    })

os.makedirs(os.path.dirname(path), exist_ok=True)
with open(path, 'w') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
    f.write('\n')
"
  echo "  hooks.json にQA hookを登録しました"
fi

echo "✅ インストール完了！Claude Codeを再起動してください。"
echo "   資料作成は「資料作って」と話しかけるだけでOKです。"
