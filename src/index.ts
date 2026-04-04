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
**以下のデザインルールに必ず従うこと。リソースを読まなくてもこのINSTRUCTIONSだけで正しいスライドを作れる。**

## いつ使うか

ユーザーが「資料作って」「スライド作って」「プレゼン作って」「提案書作って」「deck作って」と言ったら、
またはPPTXファイルに言及したら、このサーバーのツールを使って作業してください。
**HTML出力は禁止。必ずPPTXファイルを生成すること。**

---

## SlideKitデザインシステム（必須）

### カラーパレット（60-30-10ルール）
| 役割 | HEX | 用途 |
|---|---|---|
| 背景(60%) | F5F5F5 | 全スライド統一 |
| テキスト(30%) | 333333 | 本文 |
| タイトル | 222222 | スライドタイトル |
| 補足 | 666666 | 副テキスト |
| ミュート | AAAAAA | ページ番号 |
| アクセント(10%) | EF4823 | 見出し・KeyMsg・アクセントバー |
| セカンダリ | FCBF17 | YellowLine |
| KeyMsg背景 | FFF5F0 | KeyMsgBar背景 |
| セパレーター | EEEEEE | 細い区切り線 |
| デバイダー | DDDDDD | 太い縦区切り線 |

### フォント
全テキスト統一: Hiragino Kaku Gothic Pro W3

| 要素 | サイズ(pt) | bold | 色 |
|---|---|---|---|
| スライドタイトル | 22 | Yes | 222222 |
| ヒーローテキスト | 28-40 | Yes | EF4823 |
| セクション見出し | 16 | Yes | EF4823 |
| 本文 | 14 | No | 333333 |
| キーメッセージ | 18 | Yes | EF4823 |
| ページ番号 | 9 | No | AAAAAA |
| 副テキスト | 12 | No | 666666 |

### 共通コンポーネント（固定座標）
スライドサイズ: 10"×5.625"（16:9）

| 要素 | x(inch) | y(inch) | w | h | 備考 |
|---|---|---|---|---|---|
| Title | 0.5 | 0.39 | 9.0 | 0.45 | 22pt bold, 中央揃え |
| RedLine | 0.5 | 0.857 | 4.25 | 0.035 | fill=EF4823 |
| YellowLine | 4.75 | 0.857 | 4.75 | 0.035 | fill=FCBF17 |
| KeyMsgBg | 0.5 | 4.837 | 9.0 | 0.4 | roundRect, fill=FFF5F0 |
| KeyMsgText | 0.5 | 4.837 | 9.0 | 0.4 | 18pt bold, EF4823, 中央揃え, **28全角文字以内** |
| PageNum | 9.2 | 5.337 | 0.5 | 0.25 | 9pt, AAAAAA, 右揃え |

**本体コンテンツ領域**: y=0.893〜4.837 (高さ3.944")

### PptxGenJSヘルパーテンプレート
\`\`\`javascript
const C = { bg:"F5F5F5", title:"222222", body:"333333", sub:"666666", muted:"AAAAAA",
  primary:"EF4823", secondary:"FCBF17", kmBg:"FFF5F0", sep:"EEEEEE", divider:"DDDDDD", white:"FFFFFF" };
const FONT = "Hiragino Kaku Gothic Pro W3";
const SW = 10, SH = 5.625;
const HDR = { titleX:0.5, titleY:0.39, titleW:9.0, titleH:0.45,
  redLineX:0.5, redLineY:0.857, redLineW:4.25, redLineH:0.035,
  yellowLineX:4.75, yellowLineY:0.857, yellowLineW:4.75, yellowLineH:0.035 };
const KM = { x:0.5, y:4.837, w:9.0, h:0.4 };
const PN = { x:9.2, y:5.337, w:0.5, h:0.25 };
const BODY_TOP = 0.893, BODY_BOT = 4.837, BODY_H = BODY_BOT - BODY_TOP;
function centerY(contentH) { return BODY_TOP + (BODY_H - contentH) / 2; }
\`\`\`

---

## レイアウトルール

### 縦中央配置（全スライド必須・例外なし）
- **コンテンツスライド**: ヘッダー下端(0.893")〜KeyMsg上端(4.837")の領域で本体を縦中央
  - 計算: ideal_top = 0.893 + (3.944 - contentH) / 2
- **セクション/タイトルスライド**: スライド全体(5.625")で縦中央
  - 計算: ideal_top = (5.625 - contentH) / 2
- **CTA/エンドスライド**: 縦中央 AND 左右中央。固定値のx/y決め打ち禁止

### テキストボックスのサイズ
- 幅はテキスト量に合わせる（広すぎると右に空白ができる）
- 高さ(cy)は実際のテキスト量に合わせる。巨大な空きボックス禁止
- 日本語16ptで1行≈25-28文字、14ptで1行≈41文字(w=8")

### 要素間スペーシング
- グループ内: 0.15"
- グループ間: 0.3"
- セクション間: 0.6"
- gap=0は禁止

### マージン
- 外側マージン: 0.5"（全辺統一）
- コンテンツ領域: x=0.5〜9.5, y=0.5〜5.125

---

## レイアウトパターン

### コンテンツスライド(Type C)
大見出し+補足 / 左右2カラム / 3カラムグリッド / 番号付きリスト / 定義ブロック / Before→After / 2×2グリッド / 因果フロー / プロセスフロー(3ステップまで) / 番号付き縦リスト(4ステップ以上) / A/B選択肢

### 横並びプロセスフローは3ステップが上限
4ステップ以上 → 番号付き縦リストに切り替え。横並び4つは1ステップ約1.7"しかなく崩壊する。

### 矢印の方向
横並び→横矢印、縦並び→縦矢印。方向が不一致は禁止。

---

## 重要な制約

- **HTML出力禁止**: 必ずPPTXファイルを生成する
- **画像は必ず入れる**: テキストだけのプレゼンは禁止。対象企業の実物写真を使う
- **画像の縦横比は絶対に変えない**
- **提案資料は20枚以上**: 導入(2) + 分析(6-8) + 施策(8-10) + 効果(4) + クロージング(1)
- **1スライド1キーメッセージ**
- **KeyMsgは28全角文字以内**: KeyMsgBarの幅(9.0")に18pt boldで1行に収まる上限。超えたら短縮する
- **ホワイトスペース**: スライド面積の30-50%
- **バージョン管理**: _v1.pptx → _v2.pptx。上書き禁止

---

## ワークフロー

### 新規作成
1. PptxGenJS（Node.js）でスライドを生成（詳細は pptx://pptxgenjs リソース参照）
2. 上記のSlideKitデザインシステムに従う
3. pptx_thumbnail でサムネイル生成 → テキスト溢れ・切れをチェック → 問題あれば修正→再チェックをループ

### 既存PPTX編集
1. pptx_thumbnail で既存スライドを確認
2. pptx_inventory でテキスト内容を抽出
3. pptx_unpack → XMLを直接編集 → pptx_clean → pptx_pack

### 詳細ガイド（リソース）
より詳細なルールが必要な場合はリソースを参照:
- pptx://design-rules — 全体ルール・QA手順
- pptx://slidekit — OOXML座標・レイアウトパターン詳細
- pptx://pptxgenjs — PptxGenJS APIリファレンス
- pptx://rules — 破損防止チェックリスト
- pptx://editing-workflow — unpack→edit→packの詳細手順

---

## PPTX破損防止（必須）

1. presentation.xmlのスライドサイズを変更しない
2. [Content_Types].xmlに重複エントリを作らない
3. presentation.xml.relsのrIdは連番維持
4. sldIdLstとrelsのslide参照を一致させる
5. 全slideN.xml.relsが存在するか確認
6. リパック前にバリデーション実行
7. XMLエスケープ必須: & → &amp; < → &lt; > → &gt;
8. <p:sp>内でprstGeom="line"を使わない（薄い矩形で代替）
9. 不要ファイル(.bak, .tmp, .DS_Store)をzipに含めない

---

## PptxGenJS注意事項

- HEXカラーに"#"を付けない: "FF0000"が正解、"#FF0000"は破損
- 8桁HEXカラー禁止（"00000020"等）→ opacity プロパティを使う
- bullet: true を使う。Unicode "•" は二重弾丸になる
- breakLine: true でテキスト配列を改行
- オプションオブジェクトを複数呼び出しで再利用しない（内部で変更される）
- rounding: true は使わない（PowerPointエラーの原因）
- colWの合計はwと完全一致させる
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
