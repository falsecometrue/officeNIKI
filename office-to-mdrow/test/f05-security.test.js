const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const AdmZip = require("adm-zip");
const { convertDocxToMarkdown } = require("../out/converters/docxToMarkdown");
const {
  createTempOutputDirectory,
  imageMimeType,
  normalizePackageEntryPath,
  openOfficeZip,
  resolvePackagePath,
  sanitizeMermaidText,
  writeFileAtomically
} = require("../out/converters/shared");

test("UT-F05-001: Office relationship targetをパッケージ内の許可rootへ制限する", () => {
  assert.equal(resolvePackagePath("ppt/slides/slide1.xml", "../media/image1.png", "ppt/"), "ppt/media/image1.png");
  assert.equal(resolvePackagePath("xl/workbook.xml", "worksheets/sheet1.xml", "xl/"), "xl/worksheets/sheet1.xml");

  assert.throws(() => resolvePackagePath("xl/workbook.xml", "../../evil.xml", "xl/"), /escapes/);
  assert.throws(() => resolvePackagePath("word/document.xml", "http://example.test/image.png", "word/"), /Unsafe/);
  assert.throws(() => resolvePackagePath("word/document.xml", "/customXml/item1.xml", "word/"), /escapes/);
});

test("UT-F05-002: Office ZIP内の危険なentry名を拒否する", () => {
  assert.throws(() => normalizePackageEntryPath("../evil.txt"), /Unsafe Office package path/);
  assert.throws(() => normalizePackageEntryPath("word/../evil.txt"), /Unsafe Office package path/);

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "office-to-mdrow-security-"));
  const zipPath = path.join(dir, "safe.docx");
  const zip = new AdmZip();
  zip.addFile("word/document.xml", Buffer.from("<w:document/>"));
  zip.writeZip(zipPath);

  assert.doesNotThrow(() => openOfficeZip(zipPath));
});

test("UT-F05-003: 埋め込み画像は安全なraster形式だけを許可する", () => {
  assert.equal(imageMimeType("word/media/image1.png"), "image/png");
  assert.equal(imageMimeType("word/media/image1.jpg"), "image/jpeg");
  assert.equal(imageMimeType("word/media/image1.webp"), "image/webp");
  assert.equal(imageMimeType("word/media/image1.svg"), undefined);
  assert.equal(imageMimeType("word/media/image1.emf"), undefined);
});

test("UT-F05-004: Mermaidテキストは改行・コードフェンス・構文文字を無害化する", () => {
  const sanitized = sanitizeMermaidText("A```\\n%%{init: {}}\\nB-->C: done");

  assert.doesNotMatch(sanitized, /```/);
  assert.doesNotMatch(sanitized, /\n/);
  assert.doesNotMatch(sanitized, /%%/);
  assert.doesNotMatch(sanitized, /-->/);
  assert.match(sanitized, /：/);
});

test("UT-F05-005: シンボリックリンクの出力先を拒否する", () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "office-to-mdrow-security-"));
  const targetDir = path.join(dir, "target-dir");
  const linkedDir = path.join(dir, "linked-dir");
  fs.mkdirSync(targetDir);
  fs.symlinkSync(targetDir, linkedDir, "dir");

  assert.throws(() => createTempOutputDirectory(linkedDir), /symbolic link/);

  const targetFile = path.join(dir, "target.txt");
  const linkedFile = path.join(dir, "linked.txt");
  fs.writeFileSync(targetFile, "keep", "utf8");
  fs.symlinkSync(targetFile, linkedFile);

  assert.throws(() => writeFileAtomically(linkedFile, "overwrite", "utf8"), /symbolic link/);
  assert.equal(fs.readFileSync(targetFile, "utf8"), "keep");
});

test("UT-F05-006: 変換失敗時に既存出力ディレクトリを保持する", async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "office-to-mdrow-security-"));
  const docxPath = path.join(dir, "broken.docx");
  const outputDir = path.join(dir, "broken");
  fs.writeFileSync(docxPath, "not a zip", "utf8");
  fs.mkdirSync(outputDir);
  fs.writeFileSync(path.join(outputDir, "sentinel.txt"), "keep", "utf8");

  await assert.rejects(() => convertDocxToMarkdown(docxPath));
  assert.equal(fs.readFileSync(path.join(outputDir, "sentinel.txt"), "utf8"), "keep");
});
