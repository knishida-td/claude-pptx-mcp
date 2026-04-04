#!/usr/bin/env bash
# claude-pptx-mcp ワンライナーインストール
# 使い方: curl -fsSL https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main/install.sh | bash

set -euo pipefail

REPO_RAW="https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main"
SETTINGS_FILE="$HOME/.claude/settings.json"

echo "🔧 claude-pptx-mcp をインストールします..."

# settings.json がなければ作成
if [ ! -f "$SETTINGS_FILE" ]; then
  mkdir -p "$(dirname "$SETTINGS_FILE")"
  echo '{}' > "$SETTINGS_FILE"
  echo "  settings.json を新規作成しました"
fi

# MCPサーバー登録（未登録の場合のみ）
if grep -q "claude-pptx-mcp" "$SETTINGS_FILE" 2>/dev/null; then
  echo "  MCPサーバーは登録済みです"
else
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
  echo "  MCPサーバーを登録しました"
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

# ── ヘルパー: ローカルファイルを探し、なければGitHubからダウンロード ──
install_script() {
  local name="$1"       # e.g. "validate-slidekit.sh"
  local repo_path="$2"  # e.g. "scripts/validate-slidekit.sh"
  local dest="$3"       # e.g. "$HOME/.claude/scripts/validate-slidekit.sh"

  mkdir -p "$(dirname "$dest")"

  # ローカルにあれば使う（git cloneやローカル実行時）
  local local_path=""
  local script_dir=""
  script_dir="$(cd "$(dirname "$0")" 2>/dev/null && pwd)" || true
  if [ -n "$script_dir" ] && [ -f "$script_dir/$repo_path" ]; then
    local_path="$script_dir/$repo_path"
  fi

  if [ -n "$local_path" ]; then
    cp "$local_path" "$dest"
  else
    # curl | bash の場合: GitHubから直接ダウンロード
    if curl -fsSL "$REPO_RAW/$repo_path" -o "$dest" 2>/dev/null; then
      true
    else
      echo "  ⚠ $name のダウンロードに失敗しました（手動で配置してください）"
      return 1
    fi
  fi

  chmod +x "$dest"
  echo "  $name をインストールしました"
}

# validate-slidekit.sh をインストール
install_script "validate-slidekit.sh" \
  "scripts/validate-slidekit.sh" \
  "$HOME/.claude/scripts/validate-slidekit.sh"

# post-bash-pptx-qa.sh をインストール
install_script "post-bash-pptx-qa.sh" \
  "scripts/post-bash-pptx-qa.sh" \
  "$HOME/.claude/scripts/hooks/post-bash-pptx-qa.sh"

# hooks.json に PostToolUse hook を登録（既に登録済みならスキップ）
HOOKS_FILE="$HOME/.claude/hooks/hooks.json"

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
