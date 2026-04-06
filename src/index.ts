#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { execFile, execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { readFile, writeFile, unlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { z } from "zod";

const execFileAsync = promisify(execFile);

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const SCRIPTS_DIR = join(PROJECT_ROOT, "scripts");
const RESOURCES_DIR = join(PROJECT_ROOT, "resources");
const VENV_DIR = join(PROJECT_ROOT, ".venv");
const SETUP_SCRIPT = join(SCRIPTS_DIR, "setup.sh");

// Run setup if not done yet (installs python-pptx, Pillow, etc.)
if (!existsSync(join(VENV_DIR, ".setup-done"))) {
  try {
    execFileSync("bash", [SETUP_SCRIPT], {
      stdio: ["ignore", "pipe", "inherit"],
      timeout: 300_000, // 5 min max for LibreOffice install
    });
  } catch {
    // Non-fatal: tools may still work partially
    process.stderr.write(
      "[claude-pptx-mcp] WARNING: Auto-setup failed. Some tools may not work.\n"
    );
  }
}

// Use venv Python if available, otherwise fall back to system python3
const PYTHON =
  process.env.PPTX_PYTHON ??
  (existsSync(join(VENV_DIR, "bin", "python3"))
    ? join(VENV_DIR, "bin", "python3")
    : "python3");

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function runPython(
  script: string,
  args: string[],
  cwd?: string
): Promise<{ stdout: string; stderr: string }> {
  const scriptPath = join(SCRIPTS_DIR, script);
  return execFileAsync(PYTHON, [scriptPath, ...args], {
    cwd: cwd ?? process.cwd(),
    maxBuffer: 50 * 1024 * 1024,
    env: {
      ...process.env,
      PYTHONPATH: SCRIPTS_DIR,
    },
  });
}

async function runNode(
  script: string,
  args: string[]
): Promise<{ stdout: string; stderr: string }> {
  const scriptPath = join(SCRIPTS_DIR, script);
  return execFileAsync("node", [scriptPath, ...args], {
    cwd: process.cwd(),
    maxBuffer: 50 * 1024 * 1024,
    env: {
      ...process.env,
      NODE_PATH: join(PROJECT_ROOT, "node_modules"),
    },
  });
}

async function readResource(filename: string): Promise<string> {
  return readFile(join(RESOURCES_DIR, filename), "utf-8");
}

// ---------------------------------------------------------------------------
// Design rules injected into every tool response
// ---------------------------------------------------------------------------

const DESIGN_RULES_SUMMARY = `
===== SlideKit必須ルール（この指示に従うこと） =====
■ 出力: 必ずPPTXファイル。HTML出力禁止
■ カラー: 背景F5F5F5 / テキスト333333 / タイトル222222 / アクセントEF4823 / セカンダリFCBF17 / KeyMsg背景FFF5F0
■ フォント: 全テキスト Hiragino Kaku Gothic Pro W3
  タイトル22pt bold / 本文14pt / セクション見出し16pt bold(EF4823) / KeyMsg18pt bold(EF4823) / ページ番号9pt(AAAAAA)
■ 共通コンポーネント(inch): Title(0.5,0.39,9.0,0.45) / RedLine(0.5,0.857,4.25,0.035) / YellowLine(4.75,0.857,4.75,0.035) / KeyMsgBg(0.5,4.837,9.0,0.4) / PageNum(9.2,5.337,0.5,0.25)
■ 本体コンテンツ領域: y=0.893〜4.837 (高さ3.944")
■ 縦中央配置(例外なし): ideal_top = 0.893 + (3.944 - contentH) / 2
■ テキスト長制限（溢れ防止・最重要）:
  - タイトルスライド: 全角20文字以内（超過するとフォント縮小される）
  - スライドタイトル: 全角18文字以内
  - KeyMsg: 全角24文字以内（超過するとフォント縮小される）
  - KPI value: 全角6文字以内（例: 「1.2兆円」OK、「1兆2,038億円」NG→改行する）
  - 3カラムの項目: 全角10文字以内（列幅2.7"に収まる長さ）
  - definitionのtitle: 全角30文字以内
  - numbered-listは4項目まで（5項目以上はdescriptionが非表示になる）
  - definitionは3項目推奨（4項目だとdescriptionが非表示になる場合あり）
■ 禁止文字: 「---」「===」「***」等の装飾線は使わない。区切りには「 - 」を使う
■ 横並びプロセスフロー: 3ステップまで。4つ以上は縦リスト
■ 提案資料: 20枚以上
■ 画像必須: テキストだけのプレゼン禁止。実物写真を使う
■ バージョン管理: _v1.pptx → _v2.pptx。上書き禁止
■ 詳細はリソース pptx://slidekit, pptx://rules, pptx://pptxgenjs を参照
================================================`.trim();

function withDesignRules(output: string): string {
  return `${output}\n\n${DESIGN_RULES_SUMMARY}`;
}

// ---------------------------------------------------------------------------
// Server
// ---------------------------------------------------------------------------

const INSTRUCTIONS = `# PPTX資料作成MCPサーバー

## 最重要ルール

**新規スライド作成は必ず pptx_generate ツールを使うこと。**
PptxGenJSコードを自分で書くな。デザインはサーバーが制御する。

pptx_generate はJSON形式のスライド定義を受け取り、SlideKitデザインシステムに準拠したPPTXを自動生成する。
色・フォント・座標・レイアウトはすべてサーバー側でハードコードされており、変更不可。

## いつ使うか

「資料作って」「スライド作って」「プレゼン作って」「提案書作って」「deck作って」→ pptx_generate を使う。
**HTML出力は禁止。必ずPPTXファイルを生成すること。**

## pptx_generate の使い方

### JSON構造
\`\`\`json
{
  "meta": { "title": "提案書タイトル", "author": "作成者", "client": "クライアント名" },
  "slides": [
    { "type": "title", "title": "メインタイトル", "subtitle": "サブタイトル", "date": "2026年4月", "author": "blends inc." },
    { "type": "agenda", "title": "目次", "content": { "items": ["現状分析", "施策提案", "実行計画"] } },
    { "type": "section", "number": "01", "title": "現状分析" },
    {
      "type": "content",
      "title": "スライドタイトル",
      "layout": "レイアウト名",
      "content": {
        "...": "...",
        "images": [
          {
            "path": "/abs/path/to/screenshot.png",
            "x": 5.6,
            "y": 1.4,
            "w": 3.2,
            "h": 2.4,
            "altText": "製品スクリーンショット"
          }
        ]
      },
      "keyMessage": "28文字以内のメッセージ"
    },
    { "type": "cta", "title": "次のステップ", "items": [{ "label": "ラベル", "detail": "詳細" }] }
  ]
}
\`\`\`

\`images\` は \`slide.images\` または \`content.images\` に指定できる。各画像は \`{ path, x, y, w, h, altText? }\` 形式で渡し、生成側で縦横比を維持して配置する。

### 利用可能なレイアウト（type: "content" 用）
| layout | 用途 | content構造 |
|---|---|---|
| bigtext | インパクト数値・主張 | { heading: "9,800億円", subtext: "補足説明" } |
| two-column | 左右比較 | { left: { title, items: [] }, right: { title, items: [] } } |
| three-column | 3つの並列項目 | { columns: [{ title, items: [] }, ...] } |
| numbered-list | 番号付き手順・要点 | { items: [{ title, description }, ...] } |
| definition | 定義・解説 | { items: [{ title, description }, ...] } |
| before-after | 変化の表現 | { before: { title, items: [] }, after: { title, items: [] } } |
| grid-2x2 | 2×2マトリクス | { cells: [{ title, description }, ...] } |
| process-flow | 工程表示（3ステップまで） | { steps: [{ title, description }, ...] } |
| vertical-steps | 4ステップ以上の工程 | { items: [{ title, description }, ...] } |
| kpi | KPI/数値ハイライト | { metrics: [{ value, label, sub }, ...] } |
| table | 表形式データ | { headers: [], rows: [[...], ...] } |
| ab-choice | A/B選択提示 | { optionA: { title, subtitle, description }, optionB: { ... } } |
| bullets | シンプル箇条書き | { items: ["項目1", "項目2", ...] } |
| timeline | スケジュール | { phases: [{ label, title, width: 0.0-1.0 }, ...] } |

### 提案資料の構成（20枚以上必須）
1. title（1枚）
2. agenda（1枚）
3. section + content × N（現状分析: 6-8枚）
4. section + content × N（施策提案: 8-10枚）
5. section + content × N（効果・実行計画: 4枚）
6. cta（1枚）

## ワークフロー

### 新規作成（提案資料）

**Agent 2本、主要Bash 2回で完結させること。**

\`\`\`
Step 1: リサーチ Agent（sonnet、1本）
  - 一次リサーチ + 公式サイトでのファクトチェック + 商品画像DL を1本に統合
  - 分割しない

Step 2: 構成mdファイル作成
  - リサーチ結果をもとに /tmp/xxx-proposal/outline.md を作成
  - 以下の形式で全スライドの構成を定義してからコードに進む:

    ## P1: タイトル
    - type: title
    - タイトル: 〇〇株式会社御中 事業改善提案書
    - サブタイトル: ...
    - 日付: 2026年4月

    ## P2: アジェンダ
    - type: agenda / layout: numbered-list
    - items: 6項目
    - keyMsg: データに基づく分析から具体施策まで一貫したご提案

    ## P3: 貴社の理解
    - type: content / layout: definition
    - 画像: あり（右側に商品写真）
    - blocks: 企業概要 / ブランド / 主力商品 / 事業構成
    - keyMsg: ...

    （全スライド分）

  - 構成mdで確認すべきこと:
    - 20枚以上あるか
    - レイアウトが連続で同じパターンになっていないか
    - 廃盤商品を「ある」と書いていないか、既存施策を「提案」にしていないか
    - 各提案にKPI数値が付いているか
    - LTV改善施策が含まれているか

Step 3: PPTX生成（1チェーンBash）
  - validate-slidekit.sh → node generate.js → soffice PDF変換 → pdftoppm 画像化
  - 全て1つのBashコマンドチェーンで実行

Step 4: QA Agent（sonnet、1本）
  - 全スライド画像を1本のsonnetエージェントに渡す
  - opusは不要、haikuは折り返し検出が弱いのでsonnet

Step 5: 修正 → 再生成 → open
\`\`\`

### 新規作成（シンプル）
1. pptx_generate でPPTXを生成
2. pptx_thumbnail でサムネイル生成 → テキスト溢れ・切れをチェック
3. 問題あれば JSON を修正して再生成 → 再チェックをループ

### 既存PPTX編集
1. pptx_thumbnail で既存スライドを確認
2. pptx_inventory でテキスト内容を抽出
3. pptx_replace_text で一括置換、または pptx_unpack → XML編集 → pptx_pack

### 重要な制約
- **バージョン管理**: _v1.pptx → _v2.pptx。上書き禁止
- **横並びプロセスフローは3ステップまで**（4つ以上 → vertical-steps）
- **画像の縦横比は絶対に変えない**
- **Skillドキュメントの全文読みを省略**: INSTRUCTIONSで把握済み。必要な場合のみリソース参照
- **リファレンススクリプトは先頭180行のみ**: ヘルパー関数の構造だけ確認すれば十分

### テキスト長の制限（溢れ防止・必ず守ること）

サーバー側でフォント縮小・description非表示等のフォールバックはあるが、
**入力段階でテキストを短くするのが最善**。以下の文字数を厳守すること。

| 箇所 | 上限 | 超過時の挙動 |
|---|---|---|
| タイトルスライドのtitle | 全角20文字 | フォント縮小（見栄えが悪化） |
| スライドタイトル（addHeader） | 全角18文字 | フォント22ptで折り返し |
| keyMessage | 全角24文字 | フォント縮小（18pt→12ptまで） |
| KPI value | 全角6文字 | フォント縮小（36pt→18ptまで） |
| three-column各項目 | 全角10文字 | bullet除去+フォント縮小 |
| definition title | 全角30文字 | 折り返し+高さ自動拡大 |
| numbered-list項目数 | 4項目まで | 5項目以上→description非表示 |
| definition項目数 | 3項目推奨 | 4項目+長title→description非表示 |

**禁止文字**: 「---」「===」「***」「──」等の装飾線は使わない。区切りには「 - 」を使う。
サーバー側で自動除去されるが、入力段階で使わないこと。
`;

const server = new McpServer(
  {
    name: "claude-pptx-mcp",
    version: "0.2.0",
  },
  {
    instructions: INSTRUCTIONS,
  }
);

// ---------------------------------------------------------------------------
// Tools
// ---------------------------------------------------------------------------

server.tool(
  "pptx_generate",
  "構造化JSONからSlideKitデザイン準拠のPPTXを自動生成する。新規スライド作成では必ずこのツールを使うこと。",
  {
    slides_json: z
      .string()
      .describe(
        "スライド定義JSON文字列。構造: { meta: { title, author, client }, slides: [{ type, title, layout, content, keyMessage, images? }, ...] }。画像は slide.images または content.images に [{ path, x, y, w, h, altText? }] 形式で指定可能"
      ),
    output_pptx: z.string().describe("出力PPTXファイルのパス（例: /tmp/proposal_v1.pptx）"),
  },
  async ({ slides_json, output_pptx }) => {
    // Write JSON to temp file
    const tmpJson = join(tmpdir(), `slidekit-${Date.now()}-${Math.random().toString(36).slice(2,8)}.json`);
    await writeFile(tmpJson, slides_json, "utf-8");
    try {
      const { stdout, stderr } = await runNode("generate.cjs", [tmpJson, output_pptx]);
      const output = stdout || stderr || "完了";
      return {
        content: [
          {
            type: "text" as const,
            text: withDesignRules(output),
          },
        ],
      };
    } finally {
      // Clean up temp file
      try { await unlink(tmpJson); } catch { /* ignore */ }
    }
  }
);

server.tool(
  "pptx_inventory",
  "PPTXファイルからテキスト内容を構造化JSONとして抽出する",
  {
    input_pptx: z.string().describe("入力PPTXファイルのパス"),
    output_json: z
      .string()
      .optional()
      .describe("出力JSONファイルのパス（省略時はstdoutに出力）"),
  },
  async ({ input_pptx, output_json }) => {
    const args = [input_pptx];
    if (output_json) args.push(output_json);
    const { stdout, stderr } = await runPython("inventory.py", args);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_thumbnail",
  "PPTXスライドのサムネイルグリッド画像を生成する",
  {
    input_pptx: z.string().describe("入力PPTXファイルのパス"),
    output_prefix: z
      .string()
      .optional()
      .describe("出力ファイルのプレフィックス（デフォルト: thumbnails）"),
    cols: z
      .number()
      .optional()
      .describe("グリッドの列数（デフォルト: 3）"),
  },
  async ({ input_pptx, output_prefix, cols }) => {
    const args = [input_pptx];
    if (output_prefix) args.push(output_prefix);
    if (cols) args.push("--cols", String(cols));
    const { stdout, stderr } = await runPython("thumbnail.py", args);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_replace_text",
  "PPTXファイルのテキストをJSONの指定に従って一括置換する",
  {
    input_pptx: z.string().describe("入力PPTXファイルのパス"),
    replacements_json: z
      .string()
      .describe("置換定義JSONファイルのパス（inventory.pyの出力形式）"),
    output_pptx: z.string().describe("出力PPTXファイルのパス"),
  },
  async ({ input_pptx, replacements_json, output_pptx }) => {
    const { stdout, stderr } = await runPython("replace.py", [
      input_pptx,
      replacements_json,
      output_pptx,
    ]);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_rearrange",
  "PPTXスライドを指定した順序で並べ替える",
  {
    input_pptx: z.string().describe("入力PPTXファイルのパス"),
    output_pptx: z.string().describe("出力PPTXファイルのパス"),
    order: z
      .string()
      .describe("スライド順序（0始まりカンマ区切り、例: 0,3,1,2）"),
  },
  async ({ input_pptx, output_pptx, order }) => {
    const { stdout, stderr } = await runPython("rearrange.py", [
      input_pptx,
      output_pptx,
      order,
    ]);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_unpack",
  "PPTXファイルをディレクトリに展開する（XML直接編集用）",
  {
    input_pptx: z.string().describe("入力PPTXファイルのパス"),
    output_dir: z.string().describe("展開先ディレクトリのパス"),
  },
  async ({ input_pptx, output_dir }) => {
    const { stdout, stderr } = await runPython("office/unpack.py", [
      input_pptx,
      output_dir,
    ]);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_pack",
  "展開したディレクトリをPPTXファイルにパックする（バリデーション付き）",
  {
    input_dir: z.string().describe("展開済みディレクトリのパス"),
    output_pptx: z.string().describe("出力PPTXファイルのパス"),
    original_pptx: z
      .string()
      .optional()
      .describe("元のPPTXファイル（バリデーション用）"),
    validate: z
      .boolean()
      .optional()
      .describe("バリデーションを実行するか（デフォルト: true）"),
  },
  async ({ input_dir, output_pptx, original_pptx, validate }) => {
    const args = [input_dir, output_pptx];
    if (original_pptx) args.push("--original", original_pptx);
    if (validate === false) args.push("--validate", "false");
    const { stdout, stderr } = await runPython("office/pack.py", args);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_add_slide",
  "展開済みPPTXにスライドを追加・複製する",
  {
    unpacked_dir: z.string().describe("展開済みディレクトリのパス"),
    source: z
      .string()
      .describe(
        "コピー元（slide2.xml=複製, slideLayout2.xml=レイアウトから作成）"
      ),
  },
  async ({ unpacked_dir, source }) => {
    const { stdout, stderr } = await runPython("add_slide.py", [
      unpacked_dir,
      source,
    ]);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_add_image",
  "展開済みPPTXのスライドに画像を追加する",
  {
    unpacked_dir: z.string().describe("展開済みディレクトリのパス"),
    slide_num: z.number().describe("スライド番号"),
    image_path: z.string().describe("画像ファイルのパス"),
    x: z.number().optional().describe("X位置（EMU、デフォルト: 457200）"),
    y: z.number().optional().describe("Y位置（EMU、デフォルト: 1200000）"),
    cx: z.number().optional().describe("幅（EMU、省略時は自動）"),
    cy: z.number().optional().describe("高さ（EMU、省略時は自動）"),
    max_cx: z.number().optional().describe("最大幅（EMU）"),
    max_cy: z.number().optional().describe("最大高さ（EMU）"),
    name: z.string().optional().describe("シェイプ名"),
    round: z.number().optional().describe("角丸半径（EMU、0=角丸なし）"),
  },
  async ({ unpacked_dir, slide_num, image_path, x, y, cx, cy, max_cx, max_cy, name, round }) => {
    const args = [unpacked_dir, String(slide_num), image_path];
    if (x !== undefined) args.push("--x", String(x));
    if (y !== undefined) args.push("--y", String(y));
    if (cx !== undefined) args.push("--cx", String(cx));
    if (cy !== undefined) args.push("--cy", String(cy));
    if (max_cx !== undefined) args.push("--max-cx", String(max_cx));
    if (max_cy !== undefined) args.push("--max-cy", String(max_cy));
    if (name !== undefined) args.push("--name", name);
    if (round !== undefined) args.push("--round", String(round));
    const { stdout, stderr } = await runPython("add_image.py", args);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_clean",
  "展開済みPPTXから未参照ファイルを削除する",
  {
    unpacked_dir: z.string().describe("展開済みディレクトリのパス"),
  },
  async ({ unpacked_dir }) => {
    const { stdout, stderr } = await runPython("clean.py", [unpacked_dir]);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

server.tool(
  "pptx_validate",
  "展開済みPPTXのスキーマバリデーションを実行する",
  {
    unpacked_dir: z.string().describe("展開済みディレクトリのパス"),
    original_pptx: z
      .string()
      .optional()
      .describe("元のPPTXファイル（差分バリデーション用）"),
  },
  async ({ unpacked_dir, original_pptx }) => {
    const args = [unpacked_dir];
    if (original_pptx) args.push("--original", original_pptx);
    const { stdout, stderr } = await runPython("office/validate.py", args);
    return {
      content: [
        { type: "text" as const, text: withDesignRules(stdout || stderr || "完了") },
      ],
    };
  }
);

// ---------------------------------------------------------------------------
// Resources — design rules & guides
// ---------------------------------------------------------------------------

const RESOURCE_FILES: Array<{
  uri: string;
  name: string;
  description: string;
  file: string;
}> = [
  {
    uri: "pptx://design-rules",
    name: "PPTX Design Rules",
    description: "PPTXスキルの全体ルール（画像必須、バージョン管理、デザインアイデア等）",
    file: "SKILL.md",
  },
  {
    uri: "pptx://slidekit",
    name: "SlideKit Design System",
    description: "SlideKitデザインシステム（カラー、フォント、コンポーネント配置、レイアウトパターン）",
    file: "slidekit.md",
  },
  {
    uri: "pptx://editing-workflow",
    name: "Editing Workflow",
    description: "既存PPTXの編集ワークフロー（unpack→edit→pack）",
    file: "editing.md",
  },
  {
    uri: "pptx://pptxgenjs",
    name: "PptxGenJS Guide",
    description: "PptxGenJSによるゼロからのスライド作成ガイド",
    file: "pptxgenjs.md",
  },
  {
    uri: "pptx://japanese-market-rules",
    name: "Japanese Market Rules",
    description: "日本市場向けプレゼン資料ルール（20枚以上、構成テンプレート等）",
    file: "japanese-market-rules.md",
  },
  {
    uri: "pptx://rules",
    name: "PPTX Production Rules",
    description: "資料作成の必須ルール（デザインバランス、垂直中央配置、画像扱い、破損防止チェックリスト等）",
    file: "RULES.md",
  },
  {
    uri: "pptx://html2pptx",
    name: "HTML to PPTX Guide",
    description: "HTML→PPTX変換ガイド（html2pptx.jsの使い方）",
    file: "html2pptx.md",
  },
  {
    uri: "pptx://ooxml",
    name: "OOXML Reference",
    description: "OOXML直接操作リファレンス（XML構造、名前空間、要素一覧）",
    file: "ooxml.md",
  },
];

for (const res of RESOURCE_FILES) {
  server.resource(res.uri, res.uri, async () => {
    const content = await readResource(res.file);
    return {
      contents: [
        {
          uri: res.uri,
          mimeType: "text/markdown",
          text: content,
        },
      ],
    };
  });
}

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error("Failed to start MCP server:", err);
  process.exit(1);
});
