const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const { convertDocxToMarkdown } = require("../out/converters/docxToMarkdown");
const { copyFixture, unitTestDataDir } = require("./helpers");

const testDataDir = unitTestDataDir("UT_F02_word→md変換機能");

function copyTestDocument(name) {
  return copyFixture(testDataDir, name, "office-to-mdrow-f02-");
}

async function convertFixture(name) {
  const docxPath = copyTestDocument(name);
  const markdownPath = await convertDocxToMarkdown(docxPath);
  const markdown = fs.readFileSync(markdownPath, "utf8");
  const resourceDir = path.join(path.dirname(docxPath), "resources");
  const resources = fs.existsSync(resourceDir)
    ? fs.readdirSync(resourceDir).filter((entry) => fs.statSync(path.join(resourceDir, entry)).isFile())
    : [];
  return { docxPath, markdownPath, markdown, resourceDir, resources };
}

test("UT-F02-001: Markdownを生成し、UTF-8で読み込める", async () => {
  const { markdownPath, markdown } = await convertFixture("UT-F02-001.docx");

  assert.equal(path.extname(markdownPath), ".md");
  assert.equal(fs.existsSync(markdownPath), true);
  assert.ok(markdown.length > 0);
  assert.match(markdown, /新見積管理システム/);
});

test("UT-F02-003: 見出しをMarkdown見出しとして出力する", async () => {
  const { markdown } = await convertFixture("UT-F02-003_004.docx");

  assert.match(markdown, /^# .+/m);
  assert.match(markdown, /^## .+/m);
  assert.match(markdown, /^### .+/m);
  assert.match(markdown, /^#### .+/m);
});

test("UT-F02-004: 通常段落を本文として出力し、段落間に空行を入れる", async () => {
  const { markdown } = await convertFixture("UT-F02-003_004.docx");

  assert.match(markdown, /Update\/create Word and PowerPoint content/);
  assert.match(markdown, /To use this document:/);
  assert.match(markdown, /\S+\n\n\S+/);
});

test("UT-F02-012: 複数画像をresourcesへ出力し、Markdownから相対参照する", async () => {
  const { markdown, resources } = await convertFixture("UT-F02-012.docx");

  const imageLinks = markdown.match(/!\[[^\]]*]\(resources\/[^)]+\)/g) || [];
  assert.equal(resources.length, 3);
  assert.ok(imageLinks.length >= 3, `画像リンク数: ${imageLinks.length}`);
  assert.doesNotMatch(markdown, /data:image\//);
});

test("UT-F02-018: Mermaid候補または図形由来テキストをMarkdownへ出力する", async () => {
  const { markdown } = await convertFixture("UT-F02-018.docx");

  assert.ok(
    markdown.includes("```mermaid") || markdown.includes("> 図形テキスト:"),
    "Mermaidコードブロックまたは図形テキストが出力されること"
  );
  assert.match(markdown, /開始/);
  assert.match(markdown, /終了/);
});
