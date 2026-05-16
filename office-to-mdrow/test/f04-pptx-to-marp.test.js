const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const AdmZip = require("adm-zip");
const { XMLParser } = require("fast-xml-parser");
const { convertPptxToMarp } = require("../out/converters/pptxToMarp");
const { asArray, copyFixture, unitTestDataDir } = require("./helpers");

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true
});

const testDataDir = unitTestDataDir("UT_F04_powerpoint→marp変換機能");
const fixtureName = "青 シンプル ビジネス 企画書 営業 プレゼンテーション.pptx";

function copyTestPresentation() {
  return copyFixture(testDataDir, fixtureName, "office-to-mdrow-f04-");
}

function slideCount(pptxPath) {
  const zip = new AdmZip(pptxPath);
  const presentation = parser.parse(zip.readAsText("ppt/presentation.xml")).presentation;
  return asArray(presentation.sldIdLst?.sldId).length;
}

async function convertFixture() {
  const pptxPath = copyTestPresentation();
  const markdownPath = await convertPptxToMarp(pptxPath);
  const outputDir = path.dirname(markdownPath);
  const slideDir = path.join(outputDir, "slide");
  const slides = fs.readdirSync(slideDir).filter((entry) => fs.statSync(path.join(slideDir, entry)).isFile());
  return {
    pptxPath,
    markdownPath,
    markdown: fs.readFileSync(markdownPath, "utf8"),
    slideDir,
    slides
  };
}

test("UT-F04-001: PowerPointをMarpとdrawio.svgリソースへ出力する", async () => {
  const { pptxPath, markdownPath, slides } = await convertFixture();
  const expectedSlideCount = slideCount(pptxPath);

  assert.equal(path.extname(markdownPath), ".md");
  assert.equal(slides.length, expectedSlideCount);
  assert.deepEqual(slides.filter((name) => name.endsWith(".drawio.svg")).length, expectedSlideCount);
  assert.equal(slides.some((name) => name.endsWith(".drawio")), false);
  assert.equal(slides.includes("slide001.drawio.svg"), true);
});

test("UT-F04-002: MarpはslideXXX.drawio.svgのみを参照する", async () => {
  const { markdown, slides } = await convertFixture();
  const refs = markdown.match(/!\[bg contain]\(slide\/slide[0-9]{3}\.drawio\.svg\)/g) || [];

  assert.match(markdown, /^marp: true$/m);
  assert.equal(refs.length, slides.length);
});

test("UT-F04-003: drawio.svgにDraw.io編集データを埋め込む", async () => {
  const { slideDir } = await convertFixture();
  const svg = fs.readFileSync(path.join(slideDir, "slide001.drawio.svg"), "utf8");
  const parsed = parser.parse(svg);

  assert.equal(parsed.svg.width, "1920");
  assert.equal(parsed.svg.height, "1080");
  assert.match(parsed.svg.content, /<mxfile host="office-to-mdrow">/);
  assert.match(parsed.svg.content, /<mxGraphModel/);
});
