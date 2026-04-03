#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { execFile, execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
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

async function readResource(filename: string): Promise<string> {
  return readFile(join(RESOURCES_DIR, filename), "utf-8");
}

// ---------------------------------------------------------------------------
// Server
// ---------------------------------------------------------------------------

const INSTRUCTIONS = `# PPTX資料作成MCPサーバー

このサーバーはPowerPoint資料の作成・編集を行うツールとルールを提供します。

## いつ使うか

ユーザーが「資料作って」「スライド作って」「プレゼン作って」「提案書作って」「deck作って」と言ったら、
またはPPTXファイルに言及したら、このサーバーのツールを使って作業してください。

## 資料を新規作成するワークフロー

1. **まずルールを読む**: リソース pptx://rules と pptx://slidekit を読み、デザインルールを把握する
2. **作成方法を選ぶ**:
   - ゼロから作成 → pptx://pptxgenjs を読み、PptxGenJS（Node.js）でスライドを生成
   - 既存テンプレートから作成 → pptx://editing-workflow を読み、unpack→edit→packの流れで編集
3. **SlideKitデザインシステムに従う**: カラーパレット（背景#F5F5F5、アクセント#EF4823）、フォント（Hiragino Kaku Gothic Pro W3）、コンポーネント配置は全て pptx://slidekit に定義済み
4. **提案資料は20枚以上**: 導入(2枚) + 分析(6-8枚) + 施策(8-10枚) + 効果(4枚) + クロージング(1枚)
5. **全スライドのコンテンツを縦中央に配置する**（例外なし）
6. **画像は必ず入れる**: テキストだけのプレゼンは禁止。対象企業の実物写真を使う
7. **生成後はテキスト溢れチェック**: pptx_thumbnail でサムネイルを生成し、テキストの溢れ・切れがないか確認。問題があれば修正→再チェックをループ
8. **バージョン管理**: ファイル名に _v1.pptx, _v2.pptx と版番号を付ける。上書き禁止

## 既存PPTXを編集するワークフロー

1. pptx_thumbnail で既存スライドを確認
2. pptx_inventory でテキスト内容を抽出
3. pptx_unpack → XMLを直接編集 → pptx_clean → pptx_pack
4. 破損防止ルール（pptx://rules の「PPTX破損防止チェックリスト」）を必ず守る

## 重要な制約

- 横並びプロセスフローは3ステップまで。4つ以上は縦リストに切り替え
- カラー比率: 背景60%、テキスト30%、アクセント10%
- ホワイトスペース: スライド面積の30-50%
- 1スライド1キーメッセージ
- 画像の縦横比は絶対に変えない
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
        { type: "text" as const, text: stdout || stderr || "完了" },
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
