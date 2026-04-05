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

# ── CLAUDE.md にPPTX生成ルールを追記（常に最新版に更新）──
CLAUDE_MD="$HOME/.claude/CLAUDE.md"
PPTX_MARKER="<!-- pptx-mcp-rules -->"

# 既存のpptx-mcp-rulesブロックを削除してから最新を追記
if [ -f "$CLAUDE_MD" ]; then
  # マーカー間のブロックを削除
  python3 -c "
import re
with open('$CLAUDE_MD') as f:
    content = f.read()
# Remove old one-liner rule too
content = re.sub(r'.*PptxGenJS.*PPTX.*HTML出力禁止.*pptx MCP.*\n?', '', content)
# Remove marker blocks
content = re.sub(r'$PPTX_MARKER.*?$PPTX_MARKER\n?', '', content, flags=re.DOTALL)
content = content.rstrip() + '\n'
with open('$CLAUDE_MD', 'w') as f:
    f.write(content)
" 2>/dev/null || true
fi

# 最新ルールを追記
cat >> "$CLAUDE_MD" << 'PPTX_RULES_EOF'

<!-- pptx-mcp-rules -->
## PPTX資料作成ルール

**新規スライド作成は必ず pptx MCPサーバーの pptx_generate ツールを使う。**
自分でPptxGenJSコードを書くな。デザインはサーバーが制御する。HTML出力禁止。

- pptx_generate にJSON（slides配列）を渡すだけ。色・フォント・座標はサーバーが固定
- 提案資料は20枚以上: title(1) + agenda(1) + section+content(分析6-8, 施策8-10, 効果4) + cta(1)
- KeyMsgは28全角文字以内
- 生成後は pptx_thumbnail で全スライドチェック → 問題ゼロまでループ
- バージョン管理: _v1.pptx → _v2.pptx。上書き禁止
<!-- pptx-mcp-rules -->
PPTX_RULES_EOF

echo "  CLAUDE.md にPPTX生成ルールを更新しました"

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

# ── hooks設定（~/.claude/hooks/hooks.json に登録）──
HOOKS_FILE="$HOME/.claude/hooks/hooks.json"
mkdir -p "$(dirname "$HOOKS_FILE")"

python3 -c "
import json, os

path = '$HOOKS_FILE'
if os.path.exists(path):
    with open(path) as f:
        data = json.load(f)
else:
    data = {}

hooks = data.setdefault('hooks', {})

# PostToolUse: PPTX生成後のQAリマインダー
post = hooks.setdefault('PostToolUse', [])
if not any('post-bash-pptx-qa' in json.dumps(h) for h in post):
    post.append({
        'matcher': 'mcp__.*__pptx_generate',
        'hooks': [{'type': 'command', 'command': 'bash ~/.claude/scripts/hooks/post-bash-pptx-qa.sh'}],
        'description': 'PPTX生成検出時にQAリマインダーを表示'
    })

with open(path, 'w') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
    f.write('\n')
" 2>/dev/null && echo "  hooks を設定しました" || echo "  ⚠ hooks設定に失敗（手動設定が必要な場合があります）"

echo "✅ インストール完了！Claude Codeを再起動してください。"
echo "   資料作成は「資料作って」と話しかけるだけでOKです。"
echo "   デザインはサーバーが自動制御します。"
