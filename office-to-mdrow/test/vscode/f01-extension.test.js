const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const vscode = require("vscode");
const AdmZip = require("adm-zip");
const { XMLParser } = require("fast-xml-parser");

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true
});

const projectRoot = path.resolve(__dirname, "..", "..");
const repoRoot = path.resolve(projectRoot, "..");
const testDataDir = path.join(
  repoRoot,
  "doc",
  "30.developAndTest",
  "01.unitTest",
  "UT_F01_excel→drowio変換機能",
  "testData"
);

function log(message) {
  console.log(`[vscode-test] ${message}`);
}

function asArray(value) {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function copyTestWorkbook(name) {
  const source = path.join(testDataDir, name);
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "office-to-mdrow-vscode-f01-"));
  const copied = path.join(tempDir, name);
  fs.copyFileSync(source, copied);
  return copied;
}

function workbookSheetNames(workbookPath) {
  const zip = new AdmZip(workbookPath);
  const workbook = parser.parse(zip.readAsText("xl/workbook.xml"));
  return asArray(workbook.workbook?.sheets?.sheet).map((sheet) => String(sheet.name));
}

function diagramNames(parsedDrawio) {
  return asArray(parsedDrawio.mxfile?.diagram).map((diagram) => String(diagram.name));
}

async function runCommandFixture(fileName) {
  const workbookPath = copyTestWorkbook(fileName);
  await vscode.commands.executeCommand(
    "officeToMdrow.convertExcelToDrawio",
    vscode.Uri.file(workbookPath)
  );
  const drawioPath = workbookPath.replace(/\.xlsx$/i, ".drawio");
  const xml = fs.readFileSync(drawioPath, "utf8");
  return {
    workbookPath,
    drawioPath,
    xml,
    parsed: parser.parse(xml)
  };
}

async function testCase(name, body) {
  log(`START ${name}`);
  await body();
  log(`PASS  ${name}`);
}

async function run() {
  await testCase("UT-F01-001: VS Code拡張コマンドでDraw.io XMLを生成できる", async () => {
    const { drawioPath, parsed } = await runCommandFixture("UT-F01-001_012.xlsx");

    assert.equal(path.extname(drawioPath), ".drawio");
    assert.equal(fs.existsSync(drawioPath), true);
    assert.equal(parsed.mxfile.host, "app.diagrams.net");
  });

  await testCase("UT-F01-007: VS Code拡張コマンドで図形vertexを出力できる", async () => {
    const { xml } = await runCommandFixture("UT-F01-007.xlsx");
    const shapeVertices = xml.match(/<mxCell id="shape-[^"]+"[^>]*vertex="1"/g) || [];
    const shapeValues = xml.match(/<mxCell id="shape-[^"]+"[^>]*value="[^"]+"/g) || [];

    assert.ok(shapeVertices.length >= 3, `図形vertex数: ${shapeVertices.length}`);
    assert.ok(shapeValues.length >= 1, "テキスト付き図形が1件以上出力されること");
  });

  await testCase("UT-F01-008: VS Code拡張コマンドで矢印edgeを出力できる", async () => {
    const { xml } = await runCommandFixture("UT-F01-007_008.xlsx");
    const edges = xml.match(/<mxCell id="edge-[^"]+"[^>]*edge="1"/g) || [];

    assert.ok(edges.length >= 1, `edge数: ${edges.length}`);
    assert.match(xml, /endArrow=classic/);
  });

  await testCase("UT-F01-009: VS Code拡張コマンドで画像をbase64埋め込みできる", async () => {
    const { xml } = await runCommandFixture("UT-F01-009.xlsx");
    const imageShapes = xml.match(/shape=image/g) || [];
    const embeddedImages = xml.match(/data:image\//g) || [];

    assert.ok(imageShapes.length >= 2, `画像shape数: ${imageShapes.length}`);
    assert.ok(embeddedImages.length >= 2, `base64画像数: ${embeddedImages.length}`);
  });

  await testCase("UT-F01-012: VS Code拡張コマンドで入力シート名とdiagram名が一致する", async () => {
    const { workbookPath, parsed } = await runCommandFixture("UT-F01-001_012.xlsx");

    assert.deepEqual(diagramNames(parsed), workbookSheetNames(workbookPath));
  });
}

module.exports = {
  run
};
