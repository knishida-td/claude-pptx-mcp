#!/usr/bin/env node
/**
 * react-icons を一発PNG化するスクリプト
 *
 * Claudeが「資料作って」と言われた時に、アイコン名・色・サイズを指定するだけで
 * SlideKitに挿入可能なPNG画像が出来上がる。
 *
 * Usage:
 *   node icon_gen.cjs <IconName> -o <output.png> [--color RRGGBB] [--size N]
 *
 * Examples:
 *   node icon_gen.cjs FaChartLine -o /tmp/icon.png --color EF4823 --size 256
 *   node icon_gen.cjs HiOutlineSparkles -o /tmp/spark.png --color 333333 --size 128
 *
 * 対応プレフィックス（先頭の英字でreact-icons のサブモジュールを自動判定）:
 *   Fa  → react-icons/fa  (FontAwesome 5)
 *   Fa6 → react-icons/fa6 (FontAwesome 6) — Fa6Solid/Fa6Brands等は名前一致で動く
 *   Hi  → react-icons/hi  (Heroicons v1)
 *   Hi2 → react-icons/hi2 (Heroicons v2)
 *   Md  → react-icons/md  (Material Design)
 *   Bi  → react-icons/bi  (Bootstrap Icons)
 *   Bs  → react-icons/bs  (Bootstrap)
 *   Lu  → react-icons/lu  (Lucide)
 *   Ri  → react-icons/ri  (Remix)
 *   Tb  → react-icons/tb  (Tabler)
 *   Io  → react-icons/io  (Ionicons v4)
 *   Io5 → react-icons/io5 (Ionicons v5)
 *   Pi  → react-icons/pi  (Phosphor)
 *
 * 出力(JSON, 1行): { success, output, size, color, icon }
 */

const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// プレフィックスからサブモジュール名へのマップ（長いプレフィックスから順に判定）
const PREFIX_TO_LIB = [
  ["Fa6", "fa6"],
  ["Hi2", "hi2"],
  ["Io5", "io5"],
  ["Fa", "fa"],
  ["Hi", "hi"],
  ["Md", "md"],
  ["Bi", "bi"],
  ["Bs", "bs"],
  ["Lu", "lu"],
  ["Ri", "ri"],
  ["Tb", "tb"],
  ["Io", "io"],
  ["Pi", "pi"],
  ["Gi", "gi"],
  ["Si", "si"],
  ["Ai", "ai"],
];

function getIcon(name) {
  for (const [prefix, lib] of PREFIX_TO_LIB) {
    if (name.startsWith(prefix) && name.length > prefix.length) {
      // 次の文字が大文字（CamelCase継続）であることを確認
      const nextChar = name[prefix.length];
      if (nextChar === nextChar.toUpperCase()) {
        try {
          const mod = require(`react-icons/${lib}`);
          if (mod[name]) return { icon: mod[name], lib };
        } catch (e) {
          // サブモジュール未インストール
        }
      }
    }
  }
  return null;
}

async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error(
      "Usage: icon_gen.cjs <IconName> -o <output.png> [--color RRGGBB] [--size N]"
    );
    process.exit(1);
  }

  const iconName = args[0];
  let output = "/tmp/icon.png";
  let color = "333333";
  let size = 256;

  for (let i = 1; i < args.length; i++) {
    if (args[i] === "-o" || args[i] === "--output") output = args[++i];
    else if (args[i] === "--color") color = args[++i].replace(/^#/, "");
    else if (args[i] === "--size") size = parseInt(args[++i], 10);
  }

  const found = getIcon(iconName);
  if (!found) {
    console.error(
      JSON.stringify({
        success: false,
        error: `Icon not found: ${iconName}. Check the name (CamelCase + library prefix like Fa/Hi/Md/Bi/Lu/Tb).`,
      })
    );
    process.exit(1);
  }

  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(found.icon, { color: `#${color}`, size: String(size) })
  );

  await sharp(Buffer.from(svg)).png().toFile(output);
  console.log(
    JSON.stringify({
      success: true,
      output,
      size,
      color: `#${color}`,
      icon: iconName,
      lib: found.lib,
    })
  );
}

main().catch((err) => {
  console.error(JSON.stringify({ success: false, error: err.message }));
  process.exit(1);
});
