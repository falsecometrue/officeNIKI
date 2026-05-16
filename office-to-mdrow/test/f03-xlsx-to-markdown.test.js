const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const { convertXlsxToMarkdown } = require("../out/converters/xlsxToMarkdown");
const { copyFixture, unitTestDataDir } = require("./helpers");

const testDataDir = unitTestDataDir("UT_F01_excel→drowio変換機能");

function copyTestWorkbook(name) {
  return copyFixture(testDataDir, name, "office-to-mdrow-f03-");
}

async function convertFixture(name) {
  const workbookPath = copyTestWorkbook(name);
  const markdownPath = await convertXlsxToMarkdown(workbookPath);
  const workbookName = path.parse(workbookPath).name;
  const outputDir = path.join(path.dirname(workbookPath), workbookName);
  const resourceDir = path.join(outputDir, "resources");
  const markdownFiles = fs.readdirSync(outputDir).filter((entry) => entry.endsWith(".md"));
  const resources = fs.existsSync(resourceDir)
    ? fs.readdirSync(resourceDir).filter((entry) => fs.statSync(path.join(resourceDir, entry)).isFile())
    : [];
  return { markdownPath, outputDir, resourceDir, markdownFiles, resources };
}

test("UT-F03-001: Excelを単一Markdownとresourcesへ出力する", async () => {
  const { markdownPath, outputDir, resourceDir, markdownFiles, resources } = await convertFixture("UT-F01-009.xlsx");

  assert.equal(path.dirname(markdownPath), outputDir);
  assert.deepEqual(markdownFiles.sort(), ["UT-F01-009.md"]);
  assert.equal(fs.existsSync(resourceDir), true);
  assert.deepEqual(resources.sort(), ["画像あり-1.png", "画像あり-2.png", "画像あり-3.png", "画像あり-4.png"]);
});

test("UT-F03-002: Markdownから画像リソースを相対参照する", async () => {
  const { markdownPath } = await convertFixture("UT-F01-009.xlsx");
  const markdown = fs.readFileSync(markdownPath, "utf8");

  assert.match(markdown, /^# オブジェクト$/m);
  assert.match(markdown, /^# 画像あり$/m);
  assert.match(markdown, /!\[[^\]]*]\(resources\/画像あり-1\.png\)/);
});
