const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const AdmZip = require("adm-zip");
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
  const docxName = path.parse(docxPath).name;
  const resourceDir = path.join(path.dirname(docxPath), docxName, `${docxName}resource`);
  const resources = fs.existsSync(resourceDir)
    ? fs.readdirSync(resourceDir).filter((entry) => fs.statSync(path.join(resourceDir, entry)).isFile())
    : [];
  return { docxPath, markdownPath, markdown, resourceDir, resources };
}

function writeMinimalDocx(name, entries) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "office-to-mdrow-f02-generated-"));
  const docxPath = path.join(dir, name);
  const zip = new AdmZip();
  zip.addFile("[Content_Types].xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>`));
  zip.addFile("word/_rels/document.xml.rels", Buffer.from(entries.relationships || `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`));
  zip.addFile("word/document.xml", Buffer.from(entries.document));
  for (const [entry, content] of Object.entries(entries.extra || {})) {
    zip.addFile(entry, Buffer.from(content));
  }
  zip.writeZip(docxPath);
  return docxPath;
}

async function convertGeneratedDocx(docxPath) {
  const markdownPath = await convertDocxToMarkdown(docxPath);
  return fs.readFileSync(markdownPath, "utf8");
}

test("UT-F02-001: Markdownを生成し、UTF-8で読み込める", async () => {
  const { markdownPath, markdown } = await convertFixture("UT-F02-001.docx");

  assert.equal(path.extname(markdownPath), ".md");
  assert.equal(path.basename(path.dirname(markdownPath)), "UT-F02-001");
  assert.equal(fs.existsSync(markdownPath), true);
  assert.ok(markdown.length > 0);
  assert.match(markdown, /新見積管理システム/);
});

test("UT-F02-005/006/007: テキストデザインをHTML混在Markdownへ出力する", async () => {
  const docxPath = writeMinimalDocx("styled.docx", {
    document: `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:color w:val="C00000"/><w:sz w:val="28"/></w:rPr>
        <w:t>重要</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
  });

  const markdown = await convertGeneratedDocx(docxPath);

  assert.match(markdown, /<p style="text-align:center">/);
  assert.match(markdown, /<strong><span style="color:#C00000; font-size:14pt">重要<\/span><\/strong>/);
});

test("UT-F02-022: チャートをMermaidへ変換する", async () => {
  const docxPath = writeMinimalDocx("chart.docx", {
    relationships: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>`,
    document: `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:docPr id="1" name="Example Charts"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rIdChart1"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`,
    extra: {
      "word/charts/chart1.xml": `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Example Charts</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx><c:v>売上</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`
    }
  });

  const markdown = await convertGeneratedDocx(docxPath);

  assert.match(markdown, /```mermaid/);
  assert.match(markdown, /xychart-beta/);
  assert.match(markdown, /Example Charts/);
  assert.match(markdown, /bar \[10, 40]/);
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

test("UT-F02-012: 複数画像をファイル名resourceへ出力し、Markdownから相対参照する", async () => {
  const { markdown, resources } = await convertFixture("UT-F02-012.docx");

  const imageLinks = markdown.match(/!\[[^\]]*]\(UT-F02-012resource\/[^)]+\)/g) || [];
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
