const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

function run(cmd, args, options = {}) {
  return execFileSync(cmd, args, {
    cwd: path.resolve(__dirname, ".."),
    encoding: "utf8",
    stdio: ["ignore", "pipe", "pipe"],
    ...options,
  });
}

function makeTempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), "claude-pptx-mcp-test-"));
}

function writeJson(filePath, value) {
  fs.writeFileSync(filePath, JSON.stringify(value, null, 2));
}

function writeTinyPng(filePath) {
  const base64 =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9pN96xkAAAAASUVORK5CYII=";
  fs.writeFileSync(filePath, Buffer.from(base64, "base64"));
}

function zipEntries(pptxPath) {
  const script = [
    "import sys, zipfile",
    "with zipfile.ZipFile(sys.argv[1]) as zf:",
    "    print('\\n'.join(zf.namelist()))",
  ].join("\n");
  const output = run("python3", ["-c", script, pptxPath]);
  return output.trim().split("\n").filter(Boolean);
}

function readZipText(pptxPath, entryName) {
  const script = [
    "import sys, zipfile",
    "with zipfile.ZipFile(sys.argv[1]) as zf:",
    "    print(zf.read(sys.argv[2]).decode('utf-8'))",
  ].join("\n");
  return run("python3", ["-c", script, pptxPath, entryName]);
}

function testEmbedsImages() {
  const tempDir = makeTempDir();
  const inputJson = path.join(tempDir, "image-slide.json");
  const outputPptx = path.join(tempDir, "image-slide.pptx");
  const imagePath = path.join(tempDir, "tiny.png");

  writeTinyPng(imagePath);
  writeJson(inputJson, {
    meta: { title: "Image Deck", author: "test" },
    slides: [
      {
        type: "content",
        title: "Product Screenshot",
        layout: "bullets",
        content: {
          items: ["point A", "point B"],
          images: [
            {
              path: imagePath,
              x: 5.6,
              y: 1.4,
              w: 3.2,
              h: 2.4,
              altText: "tiny screenshot",
            },
          ],
        },
      },
    ],
  });

  run("node", ["scripts/generate.cjs", inputJson, outputPptx]);

  const entries = zipEntries(outputPptx);
  const relsXml = readZipText(outputPptx, "ppt/slides/_rels/slide1.xml.rels");
  const slideXml = readZipText(outputPptx, "ppt/slides/slide1.xml");
  assert(
    entries.some((entry) => /^ppt\/media\/.+/.test(entry)),
    "expected generated PPTX to include an actual image file under ppt/media/"
  );
  assert(
    /relationships\/image/.test(relsXml),
    "expected slide relationships to include an image relationship"
  );
  assert(/<p:pic[ >]/.test(slideXml), "expected slide XML to contain a picture element");
}

function testProcessFlowPreservesAllSteps() {
  const tempDir = makeTempDir();
  const inputJson = path.join(tempDir, "process-flow.json");
  const outputPptx = path.join(tempDir, "process-flow.pptx");

  writeJson(inputJson, {
    meta: { title: "Process Flow", author: "test" },
    slides: [
      {
        type: "content",
        title: "Flow",
        layout: "process-flow",
        content: {
          steps: [
            { title: "Discover", description: "step 1" },
            { title: "Design", description: "step 2" },
            { title: "Build", description: "step 3" },
            { title: "Launch", description: "step 4" },
          ],
        },
      },
    ],
  });

  run("node", ["scripts/generate.cjs", inputJson, outputPptx]);

  const slideXml = readZipText(outputPptx, "ppt/slides/slide1.xml");
  ["Discover", "Design", "Build", "Launch"].forEach((label) => {
    assert(slideXml.includes(label), `expected slide XML to include step: ${label}`);
  });
}

function testInstallHookMatchesMcpToolName() {
  const installSh = fs.readFileSync(path.resolve(__dirname, "..", "install.sh"), "utf8");
  assert(
    /mcp__.*pptx_generate/.test(installSh),
    "expected install hook matcher to target the MCP tool name, not only bare pptx_generate"
  );
}

function main() {
  const tests = [
    ["embeds image media for generated slides", testEmbedsImages],
    ["preserves 4-step process flows", testProcessFlowPreservesAllSteps],
    ["matches MCP tool name in install hook", testInstallHookMatchesMcpToolName],
  ];

  for (const [name, fn] of tests) {
    try {
      fn();
      process.stdout.write(`PASS ${name}\n`);
    } catch (error) {
      process.stderr.write(`FAIL ${name}\n${error.stack}\n`);
      process.exit(1);
    }
  }
}

main();
