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
  const sheetsDir = path.join(outputDir, "sheets");
  const resourceDir = path.join(outputDir, "resources");
  const markdownFiles = fs.readdirSync(outputDir).filter((entry) => entry.endsWith(".md"));
  const sheetMarkdownFiles = fs.existsSync(sheetsDir)
    ? fs.readdirSync(sheetsDir).filter((entry) => entry.endsWith(".md"))
    : [];
  const resources = fs.existsSync(resourceDir)
    ? fs.readdirSync(resourceDir).filter((entry) => fs.statSync(path.join(resourceDir, entry)).isFile())
    : [];
  return { markdownPath, outputDir, sheetsDir, resourceDir, markdownFiles, sheetMarkdownFiles, resources };
}

test("UT-F03-001: Excelをインデックス、シート別Markdown、resourcesへ出力する", async () => {
  const { markdownPath, outputDir, sheetsDir, resourceDir, markdownFiles, sheetMarkdownFiles, resources } = await convertFixture("UT-F01-009.xlsx");

  assert.equal(path.dirname(markdownPath), outputDir);
  assert.deepEqual(markdownFiles.sort(), ["UT-F01-009.md"]);
  assert.equal(fs.existsSync(sheetsDir), true);
  assert.deepEqual(sheetMarkdownFiles.sort(), ["01_オブジェクト.md", "02_画像あり.md"]);
  assert.equal(fs.existsSync(resourceDir), true);
  assert.deepEqual(resources.sort(), ["02_画像あり-1.png", "02_画像あり-2.png", "02_画像あり-3.png", "02_画像あり-4.png"]);
});

test("UT-F03-002: Markdownから画像リソースを相対参照する", async () => {
  const { markdownPath, sheetsDir } = await convertFixture("UT-F01-009.xlsx");
  const indexMarkdown = fs.readFileSync(markdownPath, "utf8");
  const objectMarkdown = fs.readFileSync(path.join(sheetsDir, "01_オブジェクト.md"), "utf8");
  const imageMarkdown = fs.readFileSync(path.join(sheetsDir, "02_画像あり.md"), "utf8");

  assert.match(indexMarkdown, /^# UT-F01-009$/m);
  assert.match(indexMarkdown, /- \[オブジェクト\]\(sheets\/01_%E3%82%AA%E3%83%96%E3%82%B8%E3%82%A7%E3%82%AF%E3%83%88\.md\)/);
  assert.match(indexMarkdown, /- \[画像あり\]\(sheets\/02_%E7%94%BB%E5%83%8F%E3%81%82%E3%82%8A\.md\)/);
  assert.match(objectMarkdown, /^# オブジェクト$/m);
  assert.match(imageMarkdown, /^# 画像あり$/m);
  assert.match(imageMarkdown, /<table style="border-collapse:collapse;table-layout:fixed;">/);
  assert.match(imageMarkdown, /<td style="[^"]*text-align:[^"]*">/);
  assert.match(imageMarkdown, /<img src="\.\.\/resources\/02_画像あり-1\.png" alt="[^"]*" width="\d+" height="\d+"(?: style="[^"]*")?>/);
});
