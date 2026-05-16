import * as fs from "fs";
import * as path from "path";
import AdmZip = require("adm-zip");
import { readBuffer } from "./shared";
import { buildXlsxIntermediate, XlsxDict } from "./xlsxToDrawio";

type Dict = XlsxDict;

type ResourceContext = {
  zip: AdmZip;
  resourceDir: string;
  linkBaseDir: string;
};

type MergeRange = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  rowspan: number;
  colspan: number;
};

const DEFAULT_COL_WIDTH = 12.63;
const DEFAULT_ROW_HEIGHT_PT = 15.75;
const SHEET_ORIGIN_X = 40;
const SHEET_ORIGIN_Y = 40;

type RenderedImage = {
  key: string;
  row: number;
  col: number;
  html: string;
};

type TableBounds = {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
};

export async function convertXlsxToMarkdown(sourcePath: string): Promise<string> {
  const sourceName = path.parse(sourcePath).name;
  const outputDir = path.join(path.dirname(sourcePath), sourceName);
  const indexPath = path.join(outputDir, `${sourceName}.md`);
  const sheetsDir = path.join(outputDir, "sheets");
  const resourceDir = path.join(outputDir, "resources");
  fs.rmSync(resourceDir, { recursive: true, force: true });
  fs.rmSync(sheetsDir, { recursive: true, force: true });
  fs.mkdirSync(outputDir, { recursive: true });
  fs.mkdirSync(sheetsDir, { recursive: true });
  fs.mkdirSync(resourceDir, { recursive: true });

  const intermediate = buildXlsxIntermediate(sourcePath);
  const ctx: ResourceContext = {
    zip: new AdmZip(sourcePath),
    resourceDir,
    linkBaseDir: sheetsDir
  };

  const sheets = intermediate.sheets || [];
  const sheetFiles = sheetFileNames(sheets);
  const indexLines: string[] = [`# ${escapeMarkdownHeading(sourceName)}`, "", "## Sheets", ""];
  sheets.forEach((sheet: Dict, index: number) => {
    const sheetName = String(sheet.name || "Sheet");
    const sheetFileName = sheetFiles[index];
    const sheetPath = path.join(sheetsDir, sheetFileName);
    const sheetBaseName = path.parse(sheetFileName).name;
    const body = buildSheetMarkdown(sheet, sheetBaseName, ctx).trimEnd();
    fs.writeFileSync(sheetPath, `# ${escapeMarkdownHeading(sheetName)}\n\n${body ? `${body}\n` : ""}`, "utf8");
    indexLines.push(`- [${escapeMarkdownInline(sheetName)}](sheets/${encodeMarkdownLinkPath(sheetFileName)})`);
  });

  fs.writeFileSync(indexPath, `${indexLines.join("\n")}\n`, "utf8");
  return indexPath;
}

function sanitizeFileName(value: string): string {
  return value
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+$/g, "")
    .replace(/^\.+$/g, "_")
    .slice(0, 120);
}

function sheetFileNames(sheets: Dict[]): string[] {
  const used = new Set<string>();
  return sheets.map((sheet, index) => {
    const prefix = String(index + 1).padStart(2, "0");
    const baseName = sanitizeFileName(String(sheet.name || `Sheet${index + 1}`)) || `Sheet${index + 1}`;
    let fileName = `${prefix}_${baseName}.md`;
    let suffix = 2;
    while (used.has(fileName)) {
      fileName = `${prefix}_${baseName}_${suffix}.md`;
      suffix += 1;
    }
    used.add(fileName);
    return fileName;
  });
}

function buildSheetMarkdown(sheet: Dict, sheetFileName: string, ctx: ResourceContext): string {
  const lines: string[] = [];
  const imageLinks = copySheetImages(sheet, sheetFileName, ctx);
  const tableBounds = tableBoundsForSheet(sheet);
  const tableImages = tableBounds ? imageLinks.filter((image) => image.key && isInsideTableBounds(image, tableBounds)) : [];
  const table = buildHtmlTable(sheet, tableImages, tableBounds);
  if (table) {
    lines.push(table);
  }

  const shapeTexts = (sheet.drawings || [])
    .filter((drawing: Dict) => drawing.kind === "vertex" && String(drawing.text || "").trim())
    .sort((left: Dict, right: Dict) => Number(left.y || 0) - Number(right.y || 0) || Number(left.x || 0) - Number(right.x || 0))
    .map((drawing: Dict) => renderShapeText(drawing));
  if (shapeTexts.length) {
    if (lines.length) {
      lines.push("");
    }
    lines.push(...shapeTexts);
  }

  const tableImageSet = new Set(tableImages);
  const standaloneImages = imageLinks.filter((image) => !tableImageSet.has(image));
  if (standaloneImages.length) {
    if (lines.length) {
      lines.push("");
    }
    lines.push(...standaloneImages.map((image) => image.html));
  }

  return `${lines.join("\n")}\n`;
}

function buildHtmlTable(sheet: Dict, images: RenderedImage[], bounds: TableBounds | undefined): string {
  const cells: Dict[] = sheet.cells || [];
  if (!bounds || !cells.length) {
    return "";
  }
  const styles = styleMap(sheet);
  const visibleCells = cells.filter((cell) => isVisibleCell(cell, styles));
  if (!visibleCells.length) {
    return "";
  }
  const colWidths = colWidthMap(sheet);
  const rowHeights = rowHeightMap(sheet);
  const visiblePositions = new Set(visibleCells.map((cell) => `${cell.row},${cell.col}`));
  const merges = mergeRanges(sheet).filter((merge) => visiblePositions.has(`${merge.startRow},${merge.startCol}`));
  const covered = mergedCoveredPositions(merges);
  const byMergeStart = new Map(merges.map((merge) => [`${merge.startRow},${merge.startCol}`, merge]));
  const byPosition = new Map(cells.map((cell) => [`${cell.row},${cell.col}`, cell]));
  const imagesByCell = imagesByTableCell(images, merges);

  const lines = ['<table style="border-collapse:collapse;table-layout:fixed;">'];
  lines.push("<colgroup>");
  for (let col = bounds.minCol; col <= bounds.maxCol; col += 1) {
    lines.push(`<col style="width:${colWidths[col] || excelColWidthToPx(DEFAULT_COL_WIDTH)}px">`);
  }
  lines.push("</colgroup>");
  for (let row = bounds.minRow; row <= bounds.maxRow; row += 1) {
    lines.push(`<tr style="height:${rowHeights[row] || excelRowHeightToPx(DEFAULT_ROW_HEIGHT_PT)}px">`);
    for (let col = bounds.minCol; col <= bounds.maxCol; col += 1) {
      const key = `${row},${col}`;
      if (covered.has(key)) {
        continue;
      }
      const cell = byPosition.get(key);
      const merge = byMergeStart.get(key);
      const attrs = [
        merge?.rowspan && merge.rowspan > 1 ? `rowspan="${merge.rowspan}"` : "",
        merge?.colspan && merge.colspan > 1 ? `colspan="${merge.colspan}"` : "",
        `style="${htmlAttr(cellStyle(cell, styles))}"`
      ].filter(Boolean);
      const contents = [renderCellValue(cell, styles), ...(imagesByCell.get(key) || []).map((image) => image.html)].filter(Boolean);
      lines.push(`<td ${attrs.join(" ")}>${contents.join("<br>")}</td>`);
    }
    lines.push("</tr>");
  }
  lines.push("</table>");
  return lines.join("\n");
}

function tableBoundsForSheet(sheet: Dict): TableBounds | undefined {
  const cells: Dict[] = sheet.cells || [];
  const styles = styleMap(sheet);
  const visibleCells = cells.filter((cell) => isVisibleCell(cell, styles));
  if (!visibleCells.length) {
    return undefined;
  }
  const visiblePositions = new Set(visibleCells.map((cell) => `${cell.row},${cell.col}`));
  const merges = mergeRanges(sheet).filter((merge) => visiblePositions.has(`${merge.startRow},${merge.startCol}`));
  return {
    minRow: Math.min(...visibleCells.map((cell) => Number(cell.row)), ...merges.map((merge) => merge.startRow)),
    maxRow: Math.max(...visibleCells.map((cell) => Number(cell.row)), ...merges.map((merge) => merge.endRow)),
    minCol: Math.min(...visibleCells.map((cell) => Number(cell.col)), ...merges.map((merge) => merge.startCol)),
    maxCol: Math.max(...visibleCells.map((cell) => Number(cell.col)), ...merges.map((merge) => merge.endCol))
  };
}

function isInsideTableBounds(image: RenderedImage, bounds: TableBounds): boolean {
  return (
    image.row >= bounds.minRow &&
    image.row <= bounds.maxRow &&
    image.col >= bounds.minCol &&
    image.col <= bounds.maxCol
  );
}

function imagesByTableCell(images: RenderedImage[], merges: MergeRange[]): Map<string, RenderedImage[]> {
  const byCell = new Map<string, RenderedImage[]>();
  for (const image of images) {
    if (!image.key) {
      continue;
    }
    const key = mergeStartKeyForCell(image.row, image.col, merges) || image.key;
    const list = byCell.get(key) || [];
    list.push(image);
    byCell.set(key, list);
  }
  return byCell;
}

function mergeStartKeyForCell(row: number, col: number, merges: MergeRange[]): string | undefined {
  const merge = merges.find((candidate) => (
    row >= candidate.startRow &&
    row <= candidate.endRow &&
    col >= candidate.startCol &&
    col <= candidate.endCol
  ));
  return merge ? `${merge.startRow},${merge.startCol}` : undefined;
}

function styleMap(sheet: Dict): Record<number, Dict> {
  return Object.fromEntries((sheet.styles || []).map((style: Dict) => [Number(style.style_id), style]));
}

function isVisibleCell(cell: Dict, styles: Record<number, Dict>): boolean {
  if (cellText(cell).trim()) {
    return true;
  }
  const style = styles[Number(cell.style_id || 0)] || {};
  return Boolean(style.fill_color || style.border_color || style.apply_border);
}

function excelColWidthToPx(width: number): number {
  return Math.round(width * 7 + 5);
}

function excelRowHeightToPx(heightPt: number): number {
  return Math.round((heightPt * 96) / 72);
}

function colWidthMap(sheet: Dict): Record<number, number> {
  const widths: Record<number, number> = {};
  for (const col of sheet.cols || []) {
    for (let index = Number(col.min); index <= Number(col.max); index += 1) {
      widths[index] = excelColWidthToPx(Number(col.width || DEFAULT_COL_WIDTH));
    }
  }
  return widths;
}

function rowHeightMap(sheet: Dict): Record<number, number> {
  const heights: Record<number, number> = {};
  for (const row of sheet.rows || []) {
    heights[Number(row.index)] = excelRowHeightToPx(Number(row.height || DEFAULT_ROW_HEIGHT_PT));
  }
  return heights;
}

function colNameToIndex(colName: string): number {
  let index = 0;
  for (const ch of colName.toUpperCase()) {
    index = index * 26 + ch.charCodeAt(0) - "A".charCodeAt(0) + 1;
  }
  return index;
}

function splitCellRef(ref: string): [number, number] | undefined {
  const match = ref.match(/^([A-Z]+)([0-9]+)$/);
  if (!match) {
    return undefined;
  }
  return [colNameToIndex(match[1]), Number(match[2])];
}

function mergeRanges(sheet: Dict): MergeRange[] {
  return (sheet.merged_cells || [])
    .map((mergeCell: Dict) => {
      const start = splitCellRef(String(mergeCell.start || ""));
      const end = splitCellRef(String(mergeCell.end || ""));
      if (!start || !end) {
        return undefined;
      }
      const [startCol, startRow] = start;
      const [endCol, endRow] = end;
      return {
        startRow,
        startCol,
        endRow,
        endCol,
        rowspan: endRow - startRow + 1,
        colspan: endCol - startCol + 1
      };
    })
    .filter((value: MergeRange | undefined): value is MergeRange => Boolean(value));
}

function mergedCoveredPositions(merges: MergeRange[]): Set<string> {
  const covered = new Set<string>();
  for (const merge of merges) {
    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      for (let col = merge.startCol; col <= merge.endCol; col += 1) {
        if (row !== merge.startRow || col !== merge.startCol) {
          covered.add(`${row},${col}`);
        }
      }
    }
  }
  return covered;
}

function cellStyle(cell: Dict | undefined, styles: Record<number, Dict>): string {
  const style = cell ? styles[Number(cell.style_id || 0)] || {} : {};
  const font = style.font || {};
  const declarations: Record<string, string | undefined> = {
    border: `1px solid ${style.border_color || "#D9D9D9"}`,
    background: style.fill_color || undefined,
    "text-align": cssHorizontalAlign(String(style.horizontal_alignment || "")),
    "vertical-align": cssVerticalAlign(String(style.vertical_alignment || "")),
    "font-family": font.name ? `${font.name}, sans-serif` : undefined,
    "font-size": font.size_pt ? `${font.size_pt}pt` : undefined,
    color: font.color,
    padding: "2px 4px",
    "white-space": style.wrap_text ? "pre-wrap" : "normal"
  };
  return styleAttr(declarations);
}

function cssHorizontalAlign(value: string): string {
  const map: Record<string, string> = {
    center: "center",
    right: "right",
    left: "left",
    fill: "left",
    justify: "justify",
    distributed: "justify"
  };
  return map[value] || "left";
}

function cssVerticalAlign(value: string): string {
  const map: Record<string, string> = {
    top: "top",
    center: "middle",
    bottom: "bottom",
    justify: "middle",
    distributed: "middle"
  };
  return map[value] || "middle";
}

function styleAttr(styles: Record<string, string | undefined>): string {
  return Object.entries(styles)
    .filter(([, value]) => value !== undefined && value !== "")
    .map(([key, value]) => `${key}:${value}`)
    .join("; ");
}

function renderCellValue(cell: Dict | undefined, styles: Record<number, Dict>): string {
  if (!cell) {
    return "";
  }
  const style = styles[Number(cell.style_id || 0)] || {};
  const font = style.font || {};
  let rendered = htmlEscape(cellText(cell)).replace(/\r?\n/g, "<br>");
  if (font.underline) {
    rendered = `<u>${rendered}</u>`;
  }
  if (font.bold) {
    rendered = `<strong>${rendered}</strong>`;
  }
  if (font.italic) {
    rendered = `<em>${rendered}</em>`;
  }
  return rendered;
}

function copySheetImages(sheet: Dict, sheetFileName: string, ctx: ResourceContext): RenderedImage[] {
  const images = (sheet.drawings || []).filter((drawing: Dict) => drawing.kind === "image" && drawing.path);
  const colWidths = colWidthMap(sheet);
  const rowHeights = rowHeightMap(sheet);
  return images.map((image: Dict, index: number) => {
    const sourcePath = String(image.path);
    const ext = imageExtension(sourcePath);
    const fileName = `${sheetFileName}-${index + 1}${ext}`;
    const destination = path.join(ctx.resourceDir, fileName);
    fs.writeFileSync(destination, readBuffer(ctx.zip, sourcePath));
    const relPath = path.relative(ctx.linkBaseDir, destination).split(path.sep).join("/");
    const alt = htmlAttr(String(image.name || path.parse(fileName).name));
    const width = Number(image.width || 0);
    const height = Number(image.height || 0);
    const position = sheetCellAtPoint(Number(image.x || 0), Number(image.y || 0), colWidths, rowHeights);
    const key = position ? `${position.row},${position.col}` : "";
    const style = imageCellStyle(image, position, colWidths, rowHeights);
    const styleAttrText = style ? ` style="${htmlAttr(style)}"` : "";
    const html = width > 0 && height > 0
      ? `<img src="${htmlAttr(relPath)}" alt="${alt}" width="${Math.round(width)}" height="${Math.round(height)}"${styleAttrText}>`
      : `<img src="${htmlAttr(relPath)}" alt="${alt}"${styleAttrText}>`;
    return {
      key,
      row: position?.row || 0,
      col: position?.col || 0,
      html
    };
  });
}

function sheetCellAtPoint(x: number, y: number, colWidths: Record<number, number>, rowHeights: Record<number, number>): { row: number; col: number } | undefined {
  if (!Number.isFinite(x) || !Number.isFinite(y) || x < SHEET_ORIGIN_X || y < SHEET_ORIGIN_Y) {
    return undefined;
  }
  return {
    col: axisIndexAtPoint(x - SHEET_ORIGIN_X, colWidths, DEFAULT_COL_WIDTH, excelColWidthToPx),
    row: axisIndexAtPoint(y - SHEET_ORIGIN_Y, rowHeights, DEFAULT_ROW_HEIGHT_PT, excelRowHeightToPx)
  };
}

function axisIndexAtPoint(offset: number, sizes: Record<number, number>, defaultSize: number, converter: (value: number) => number): number {
  const defaultPx = converter(defaultSize);
  let cursor = 0;
  for (let index = 1; index < 16384; index += 1) {
    const size = sizes[index] || defaultPx;
    if (offset < cursor + size) {
      return index;
    }
    cursor += size;
  }
  return 16384;
}

function imageCellStyle(image: Dict, position: { row: number; col: number } | undefined, colWidths: Record<number, number>, rowHeights: Record<number, number>): string {
  const declarations: Record<string, string | undefined> = {
    display: "block",
    "max-width": "100%",
    height: "auto"
  };
  if (position) {
    const left = Math.max(0, Math.round(Number(image.x || 0) - sheetXForCol(position.col, colWidths)));
    const top = Math.max(0, Math.round(Number(image.y || 0) - sheetYForRow(position.row, rowHeights)));
    if (left || top) {
      declarations.margin = `${top}px 0 0 ${left}px`;
    }
  }
  return styleAttr(declarations);
}

function sheetXForCol(col: number, colWidths: Record<number, number>): number {
  const defaultPx = excelColWidthToPx(DEFAULT_COL_WIDTH);
  let x = SHEET_ORIGIN_X;
  for (let index = 1; index < col; index += 1) {
    x += colWidths[index] || defaultPx;
  }
  return x;
}

function sheetYForRow(row: number, rowHeights: Record<number, number>): number {
  const defaultPx = excelRowHeightToPx(DEFAULT_ROW_HEIGHT_PT);
  let y = SHEET_ORIGIN_Y;
  for (let index = 1; index < row; index += 1) {
    y += rowHeights[index] || defaultPx;
  }
  return y;
}

function renderShapeText(drawing: Dict): string {
  const styles = styleAttr({
    "background-color": drawing.fill,
    color: drawing.stroke
  });
  const text = htmlEscape(String(drawing.text || "").trim()).replace(/\r?\n/g, "<br>");
  return styles ? `<p style="${styles}">${text}</p>` : text;
}

function imageExtension(sourcePath: string): string {
  const ext = path.extname(sourcePath).toLowerCase();
  return ext || ".png";
}

function cellText(cell: Dict): string {
  return String(cell.value ?? cell.raw_value ?? "");
}

function htmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function htmlAttr(value: string): string {
  return htmlEscape(value).replace(/"/g, "&quot;");
}

function escapeMarkdownInline(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/\]/g, "\\]");
}

function escapeMarkdownHeading(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/^#+\s*/, "");
}

function encodeMarkdownLinkPath(value: string): string {
  return value
    .split("/")
    .map((part) => encodeURIComponent(part).replace(/%20/g, "%20"))
    .join("/");
}
