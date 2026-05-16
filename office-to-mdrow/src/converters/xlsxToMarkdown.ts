import * as fs from "fs";
import * as path from "path";
import AdmZip = require("adm-zip");
import { readBuffer } from "./shared";
import { buildXlsxIntermediate, XlsxDict } from "./xlsxToDrawio";

type Dict = XlsxDict;

type ResourceContext = {
  zip: AdmZip;
  outputDir: string;
  resourceDir: string;
};

export async function convertXlsxToMarkdown(sourcePath: string): Promise<string> {
  const sourceName = path.parse(sourcePath).name;
  const outputDir = path.join(path.dirname(sourcePath), sourceName);
  const markdownPath = path.join(outputDir, `${sourceName}.md`);
  const resourceDir = path.join(outputDir, "resources");
  fs.rmSync(resourceDir, { recursive: true, force: true });
  fs.mkdirSync(resourceDir, { recursive: true });

  const intermediate = buildXlsxIntermediate(sourcePath);
  const ctx: ResourceContext = {
    zip: new AdmZip(sourcePath),
    outputDir,
    resourceDir
  };

  fs.mkdirSync(outputDir, { recursive: true });
  const lines: string[] = [];
  for (const sheet of intermediate.sheets || []) {
    if (lines.length) {
      lines.push("");
    }
    const sheetName = String(sheet.name || "Sheet");
    const resourceName = sanitizeFileName(sheetName) || "Sheet";
    lines.push(`# ${escapeMarkdownHeading(sheetName)}`, "");
    const body = buildSheetMarkdown(sheet, resourceName, ctx).trimEnd();
    if (body) {
      lines.push(body);
    }
  }

  fs.writeFileSync(markdownPath, `${lines.join("\n")}\n`, "utf8");
  return markdownPath;
}

function sanitizeFileName(value: string): string {
  return value
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+$/g, "")
    .replace(/^\.+$/g, "_")
    .slice(0, 120);
}

function buildSheetMarkdown(sheet: Dict, sheetFileName: string, ctx: ResourceContext): string {
  const lines: string[] = [];
  const table = buildMarkdownTable(sheet.cells || []);
  if (table) {
    lines.push(table);
  }

  const shapeTexts = (sheet.drawings || [])
    .filter((drawing: Dict) => drawing.kind === "vertex" && String(drawing.text || "").trim())
    .map((drawing: Dict) => String(drawing.text).trim());
  if (shapeTexts.length) {
    if (lines.length) {
      lines.push("");
    }
    lines.push(...shapeTexts.map((text: string) => `- ${escapeMarkdownInline(text)}`));
  }

  const imageLinks = copySheetImages(sheet, sheetFileName, ctx);
  if (imageLinks.length) {
    if (lines.length) {
      lines.push("");
    }
    lines.push(...imageLinks);
  }

  return `${lines.join("\n")}\n`;
}

function buildMarkdownTable(cells: Dict[]): string {
  if (!cells.length) {
    return "";
  }
  const minRow = Math.min(...cells.map((cell) => Number(cell.row)));
  const maxRow = Math.max(...cells.map((cell) => Number(cell.row)));
  const minCol = Math.min(...cells.map((cell) => Number(cell.col)));
  const maxCol = Math.max(...cells.map((cell) => Number(cell.col)));
  const byPosition = new Map(cells.map((cell) => [`${cell.row},${cell.col}`, cell]));
  const rows: string[][] = [];
  for (let row = minRow; row <= maxRow; row += 1) {
    const values: string[] = [];
    for (let col = minCol; col <= maxCol; col += 1) {
      values.push(escapeMarkdownTableCell(String(byPosition.get(`${row},${col}`)?.value || "")));
    }
    rows.push(values);
  }

  const header = rows[0];
  const separator = header.map(() => "---");
  const body = rows.slice(1);
  return [header, separator, ...body]
    .map((row) => `| ${row.join(" | ")} |`)
    .join("\n");
}

function copySheetImages(sheet: Dict, sheetFileName: string, ctx: ResourceContext): string[] {
  const images = (sheet.drawings || []).filter((drawing: Dict) => drawing.kind === "image" && drawing.path);
  return images.map((image: Dict, index: number) => {
    const sourcePath = String(image.path);
    const ext = imageExtension(sourcePath);
    const fileName = `${sheetFileName}-${index + 1}${ext}`;
    const destination = path.join(ctx.resourceDir, fileName);
    fs.writeFileSync(destination, readBuffer(ctx.zip, sourcePath));
    const relPath = path.relative(ctx.outputDir, destination).split(path.sep).join("/");
    const alt = escapeMarkdownInline(String(image.name || path.parse(fileName).name));
    return `![${alt}](${relPath})`;
  });
}

function imageExtension(sourcePath: string): string {
  const ext = path.extname(sourcePath).toLowerCase();
  return ext || ".png";
}

function escapeMarkdownTableCell(value: string): string {
  return escapeMarkdownInline(value).replace(/\r?\n/g, "<br>").replace(/\|/g, "\\|");
}

function escapeMarkdownInline(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/\]/g, "\\]");
}

function escapeMarkdownHeading(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/^#+\s*/, "");
}
