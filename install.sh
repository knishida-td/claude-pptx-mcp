#!/usr/bin/env bash
# claude-pptx-mcp ワンライナーインストール
# 使い方: curl -fsSL https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main/install.sh | bash

set -euo pipefail

# curl | bash は非対話シェルで .zshrc/.bashrc を読まないため、
# nvm/nodebrew/Homebrew 等で入れた node や claude にPATHが通らない。
# ユーザーのプロファイルを明示的にロードする。
for f in "$HOME/.zprofile" "$HOME/.zshrc" "$HOME/.bash_profile" "$HOME/.bashrc" "$HOME/.profile"; do
  [ -f "$f" ] && source "$f" 2>/dev/null || true
done

REPO_RAW="https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main"

echo "🔧 claude-pptx-mcp をインストールします..."

# npxキャッシュをクリア（旧バージョンが使われるのを防止）
echo "  npxキャッシュをクリア中..."
npm cache clean --force 2>/dev/null || true
# npx のパッケージキャッシュから claude-pptx-mcp を削除
find "$(npm config get cache 2>/dev/null || echo "$HOME/.npm")" -path "*claude-pptx-mcp*" -exec rm -rf {} + 2>/dev/null || true
# _npx キャッシュも削除
find "$HOME/.npm/_npx" -path "*claude-pptx-mcp*" -exec rm -rf {} + 2>/dev/null || true
echo "  キャッシュクリア完了"

# ~/.claude ディレクトリを確保
mkdir -p "$HOME/.claude"

# MCPサーバー登録（claude mcp add を使用 — 確実に永続化される）
# Node.js チェック
if ! command -v node &>/dev/null; then
  echo "❌ Node.js が見つかりません。先にインストールしてください:"
  echo "   https://nodejs.org/"
  exit 1
fi

# claude CLI チェック
if ! command -v claude &>/dev/null; then
  echo "❌ Claude Code CLI が見つかりません。先にインストールしてください:"
  echo "   npm install -g @anthropic-ai/claude-code"
  exit 1
fi

# 既存登録を削除してから最新版で再登録（冪等性を保証）
claude mcp remove --scope user pptx 2>/dev/null || true
if claude mcp add --scope user pptx -- npx -y "github:knishida-td/claude-pptx-mcp" 2>&1; then
  echo "  MCPサーバーを登録しました（claude mcp add --scope user）"
else
  echo "❌ MCP登録に失敗しました。Claude Codeに一度ログインしてから再実行してください。"
  exit 1
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
