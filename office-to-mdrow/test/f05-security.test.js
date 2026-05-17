const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const AdmZip = require("adm-zip");
const {
  imageMimeType,
  normalizePackageEntryPath,
  openOfficeZip,
  resolvePackagePath,
  sanitizeMermaidText
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
