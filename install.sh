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

echo "✅ インストール完了！Claude Codeを再起動してください。"
echo "   資料作成は「資料作って」と話しかけるだけでOKです。"
