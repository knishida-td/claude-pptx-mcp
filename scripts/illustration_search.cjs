#!/usr/bin/env node
/**
 * SVGイラストダウンロード + 色置換 + PNG変換ヘルパー
 *
 * undraw.co / Storyset / open-peeps など、SVG形式のフリーイラスト素材を
 * 取得して SlideKit プライマリ色に合わせ、PNG化してスライドに入れる用途を想定。
 *
 * Usage:
 *   node illustration_search.cjs <svg-url> -o output.png [--color EF4823] [--width 1200]
 *
 * Options:
 *   -o, --output    出力ファイルパス（.png または .svg）。デフォルト /tmp/illustration.png
 *   --color RRGGBB  undraw のデフォルト紫(#6C63FF)を指定色に置換
 *   --width N       PNG変換時の幅ピクセル（デフォルト 1200）
 *   --svg           PNG変換せずSVGのまま保存
 *
 * 出力(JSON, 1行): { success, output, format, width? }
 */

const fs = require("fs");
const https = require("https");
const http = require("http");
const sharp = require("sharp");

const UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 " +
  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";

function fetchText(url, redirectsLeft = 5) {
  return new Promise((resolve, reject) => {
    const client = url.startsWith("https") ? https : http;
    const req = client.get(url, { headers: { "User-Agent": UA, Accept: "image/svg+xml,*/*" } }, (res) => {
      const status = res.statusCode;
      if (status >= 300 && status < 400 && res.headers.location) {
        if (redirectsLeft <= 0) {
          reject(new Error("Too many redirects"));
          return;
        }
        res.resume();
        const next = new URL(res.headers.location, url).href;
        fetchText(next, redirectsLeft - 1).then(resolve, reject);
        return;
      }
      if (status !== 200) {
        reject(new Error(`HTTP ${status} for ${url}`));
        return;
      }
      let body = "";
      res.setEncoding("utf-8");
      res.on("data", (chunk) => (body += chunk));
      res.on("end", () => resolve(body));
    });
    req.on("error", reject);
    req.setTimeout(15000, () => req.destroy(new Error("Timeout")));
  });
}

function recolorSvg(svg, targetColor) {
  if (!targetColor.startsWith("#")) targetColor = "#" + targetColor;
  // undraw のデフォルト primary 色（紫）を置換
  return svg
    .replace(/#6c63ff/gi, targetColor)
    .replace(/(?<![#A-Z0-9])6c63ff(?![A-Z0-9])/gi, targetColor.replace("#", ""));
}

async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error(
      "Usage: illustration_search.cjs <url> -o <output> [--color RRGGBB] [--width N] [--svg]"
    );
    process.exit(1);
  }
  const url = args[0];
  let output = "/tmp/illustration.png";
  let color = null;
  let width = 1200;
  let svgOnly = false;
  for (let i = 1; i < args.length; i++) {
    if (args[i] === "-o" || args[i] === "--output") output = args[++i];
    else if (args[i] === "--color") color = args[++i];
    else if (args[i] === "--width") width = parseInt(args[++i], 10);
    else if (args[i] === "--svg") svgOnly = true;
  }

  let svg = await fetchText(url);
  if (color) svg = recolorSvg(svg, color);

  if (svgOnly || output.endsWith(".svg")) {
    fs.writeFileSync(output, svg, "utf-8");
    console.log(JSON.stringify({ success: true, output, format: "svg" }));
    return;
  }

  await sharp(Buffer.from(svg)).resize({ width }).png().toFile(output);
  console.log(JSON.stringify({ success: true, output, format: "png", width }));
}

main().catch((err) => {
  console.error(JSON.stringify({ success: false, error: err.message }));
  process.exit(1);
});
