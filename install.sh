#!/usr/bin/env bash
# claude-pptx-mcp ワンライナーインストール
# 使い方: curl -fsSL https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main/install.sh | bash

set -euo pipefail

# curl | bash は非対話シェルでシェル初期化を読まないため、
# nvm/nodebrew/Homebrew 等で入れた node や claude にPATHが通らない。
# bash から直接 .zshrc を source せず、各シェル自身にPATHだけ出力させる。
recover_path_from_shell() {
  local shell_bin="$1"
  local output=""
  local marker="__CLAUDE_PPTX_PATH__"

  [ -n "$shell_bin" ] || return 1
  [ -x "$shell_bin" ] || return 1

  output="$("$shell_bin" -lic "printf '${marker}%s\n' \"\$PATH\"" 2>/dev/null)" || return 1
  output="${output##*${marker}}"
  [ -n "$output" ] || return 1

  PATH="$output"
}

for shell_bin in "${SHELL:-}" /bin/zsh /usr/bin/zsh /bin/bash /usr/bin/bash; do
  recover_path_from_shell "$shell_bin" || true
done

REPO_RAW="https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main"
MCP_NAME="pptx"
MCP_CONFIG_PATH="$HOME/.claude.json"

read_user_mcp_config() {
  python3 - "$MCP_CONFIG_PATH" "$1" <<'PY'
import json
import os
import sys

path, name = sys.argv[1], sys.argv[2]
if not os.path.exists(path):
    sys.exit(1)

with open(path) as f:
    data = json.load(f)

server = data.get("mcpServers", {}).get(name)
if server is None:
    sys.exit(1)

print(json.dumps(server, ensure_ascii=False))
PY
}

write_user_mcp_config() {
  python3 - "$MCP_CONFIG_PATH" "$1" "$2" <<'PY'
import json
import os
import sys

path, name, server_json = sys.argv[1], sys.argv[2], sys.argv[3]
server = json.loads(server_json)

data = {}
if os.path.exists(path):
    with open(path) as f:
        data = json.load(f)

mcp_servers = data.setdefault("mcpServers", {})
mcp_servers[name] = server

with open(path, "w") as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
    f.write("\n")
PY
}

echo "🔧 claude-pptx-mcp をインストールします..."

# ~/.claude ディレクトリを確保
mkdir -p "$HOME/.claude"

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

# ── リポジトリをローカルにclone（起動時にgit pullで常に最新化）──
REPO_DIR="$HOME/.claude/mcp-servers/claude-pptx-mcp"
REPO_URL="https://github.com/knishida-td/claude-pptx-mcp.git"

if [ -d "$REPO_DIR/.git" ]; then
  echo "  既存リポジトリを更新中..."
  cd "$REPO_DIR" && git pull --ff-only -q 2>/dev/null || true
else
  echo "  リポジトリをclone中..."
  mkdir -p "$(dirname "$REPO_DIR")"
  rm -rf "$REPO_DIR"
  git clone -q "$REPO_URL" "$REPO_DIR"
  cd "$REPO_DIR"
fi

# 依存インストール + ビルド
cd "$REPO_DIR"
npm install -q 2>/dev/null || npm install
if [ ! -f dist/index.js ] || [ src/index.ts -nt dist/index.js ]; then
  npx tsc 2>/dev/null || true
fi
echo "  リポジトリセットアップ完了: $REPO_DIR"

# ── launcher.shをインストール ──
LAUNCHER="$HOME/.claude/scripts/pptx-mcp-launcher.sh"
mkdir -p "$(dirname "$LAUNCHER")"
cp "$REPO_DIR/scripts/launcher.sh" "$LAUNCHER"
chmod +x "$LAUNCHER"
echo "  ランチャーをインストールしました: $LAUNCHER"

# ── MCPサーバー登録（launcher.sh経由 — 起動時にgit pullで常に最新）──
existing_mcp_config="$(read_user_mcp_config "$MCP_NAME" 2>/dev/null || true)"
add_output=""

# 旧npx方式が残っていれば削除
claude mcp remove --scope user "$MCP_NAME" 2>/dev/null || true

if add_output="$(claude mcp add --scope user "$MCP_NAME" -- bash "$LAUNCHER" 2>&1)"; then
  echo "$add_output"
  echo "  MCPサーバーを登録しました（launcher.sh経由・自動更新）"
else
  add_status=$?
  echo "$add_output"
  if [ -n "$existing_mcp_config" ]; then
    echo "  ⚠ 再登録に失敗したため、既存のMCP設定を復元します..."
    write_user_mcp_config "$MCP_NAME" "$existing_mcp_config"
    echo "❌ MCP登録に失敗しました。既存の設定は保持しています。"
    exit 1
  else
    echo "❌ MCP登録に失敗しました。Claude Codeに一度ログインしてから再実行してください。"
    exit "$add_status"
  fi
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
自分でPptxGenJSコードを書くな。デザインはサーバーが制御する。
**Google Slides・Canva・HTML・その他のスライドツールは一切使用禁止。** 資料作成=pptx_generate一択。

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

# PreToolUse: Google Slides/Canva等のスライドツール呼び出しをブロック
pre = hooks.setdefault('PreToolUse', [])
block_matcher = 'mcp__.*__(googleslides|google_slides|canva|gamma).*'
if not any(block_matcher in json.dumps(h) for h in pre):
    pre.append({
        'matcher': block_matcher,
        'hooks': [{'type': 'command', 'command': 'echo \"❌ スライド作成はpptx MCPサーバー(pptx_generate)を使ってください。Google Slides/Canva等は禁止です。\" >&2; exit 2'}],
        'description': 'Google Slides/Canva等のスライドツール呼び出しをブロック'
    })

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
