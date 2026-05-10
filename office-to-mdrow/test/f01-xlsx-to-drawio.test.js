const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const AdmZip = require("adm-zip");
const { XMLParser } = require("fast-xml-parser");
const { convertXlsxToDrawio } = require("../out/converters/xlsxToDrawio");
const { asArray, copyFixture, unitTestDataDir } = require("./helpers");

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true
});

const testDataDir = unitTestDataDir("UT_F01_excel→drowio変換機能");

function copyTestWorkbook(name) {
  return copyFixture(testDataDir, name, "office-to-mdrow-f01-");
}

async function convertFixture(name) {
  const workbookPath = copyTestWorkbook(name);
  const drawioPath = await convertXlsxToDrawio(workbookPath);
  const xml = fs.readFileSync(drawioPath, "utf8");
  const parsed = parser.parse(xml);
  return { workbookPath, drawioPath, xml, parsed };
}

function workbookSheetNames(workbookPath) {
  const zip = new AdmZip(workbookPath);
  const workbook = parser.parse(zip.readAsText("xl/workbook.xml"));
  return asArray(workbook.workbook?.sheets?.sheet).map((sheet) => String(sheet.name));
}

function diagramNames(parsedDrawio) {
  return asArray(parsedDrawio.mxfile?.diagram).map((diagram) => String(diagram.name));
}

test("UT-F01-001: Draw.io XMLを生成し、XML parseできる", async () => {
  const { drawioPath, parsed } = await convertFixture("UT-F01-001_012.xlsx");

  assert.equal(path.extname(drawioPath), ".drawio");
  assert.equal(fs.existsSync(drawioPath), true);
  assert.equal(parsed.mxfile.host, "app.diagrams.net");
});

test("UT-F01-007: 図形がvertexとして出力される", async () => {
  const { xml } = await convertFixture("UT-F01-007.xlsx");

  const shapeVertices = xml.match(/<mxCell id="shape-[^"]+"[^>]*vertex="1"/g) || [];
  const shapeValues = xml.match(/<mxCell id="shape-[^"]+"[^>]*value="[^"]+"/g) || [];

  assert.ok(shapeVertices.length >= 3, `図形vertex数: ${shapeVertices.length}`);
  assert.ok(shapeValues.length >= 1, "テキスト付き図形が1件以上出力されること");
});

test("UT-F01-008: 矢印がedgeとして出力され、終端矢印を持つ", async () => {
  const { xml } = await convertFixture("UT-F01-007_008.xlsx");

  const edges = xml.match(/<mxCell id="edge-[^"]+"[^>]*edge="1"/g) || [];

  assert.ok(edges.length >= 1, `edge数: ${edges.length}`);
  assert.match(xml, /endArrow=classic/);
});

test("UT-F01-009: 画像がDraw.io XMLへbase64埋め込みされる", async () => {
  const { xml } = await convertFixture("UT-F01-009.xlsx");

  const imageShapes = xml.match(/shape=image/g) || [];
  const embeddedImages = xml.match(/data:image\//g) || [];

  assert.ok(imageShapes.length >= 2, `画像shape数: ${imageShapes.length}`);
  assert.ok(embeddedImages.length >= 2, `base64画像数: ${embeddedImages.length}`);
});

test("UT-F01-012: 入力シートと同数のdiagramを出力し、diagram名がシート名と一致する", async () => {
  const { workbookPath, parsed } = await convertFixture("UT-F01-001_012.xlsx");

  assert.deepEqual(diagramNames(parsed), workbookSheetNames(workbookPath));
});
