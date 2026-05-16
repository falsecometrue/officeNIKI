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

type RenderedImage = {
  x: number;
  y: number;
  html: string;
};

type TableBounds = {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
};

type TableBlock = {
  bounds: TableBounds;
  cells: Dict[];
  merges: MergeRange[];
};

type MarkdownBlock = {
  x: number;
  y: number;
  markdown: string;
};

type MermaidBlock = {
  x: number;
  y: number;
  code: string;
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
  const blocks: MarkdownBlock[] = [];
  const imageLinks = copySheetImages(sheet, sheetFileName, ctx);

  const tables = tableBlocksForSheet(sheet);
  tables.forEach((table, index) => {
    blocks.push({
      x: table.bounds.minCol * 100,
      y: table.bounds.minRow * 24,
      markdown: [`## Table ${index + 1}`, "", buildHtmlTable(table, styleMap(sheet))].join("\n")
    });
  });

  const sequenceDiagram = mermaidSequenceDiagram(sheet);
  if (sequenceDiagram) {
    blocks.push({
      x: sequenceDiagram.x,
      y: sequenceDiagram.y,
      markdown: ["## Diagram 1", "", "```mermaid", sequenceDiagram.code, "```"].join("\n")
    });
  }

  imageLinks.forEach((image, index) => {
    blocks.push({
      x: image.x,
      y: image.y,
      markdown: [`## Image ${index + 1}`, "", image.html].join("\n")
    });
  });

  blocks.sort((left, right) => left.y - right.y || left.x - right.x);
  return `${blocks.map((block) => block.markdown).join("\n\n")}\n`;
}

function buildHtmlTable(table: TableBlock, styles: Record<number, Dict>): string {
  const byPosition = new Map(table.cells.map((cell) => [`${cell.row},${cell.col}`, cell]));
  const covered = mergedCoveredPositions(table.merges);
  const byMergeStart = new Map(table.merges.map((merge) => [`${merge.startRow},${merge.startCol}`, merge]));

  const lines = ['<table>'];
  for (let row = table.bounds.minRow; row <= table.bounds.maxRow; row += 1) {
    lines.push("  <tr>");
    for (let col = table.bounds.minCol; col <= table.bounds.maxCol; col += 1) {
      const key = `${row},${col}`;
      if (covered.has(key)) {
        continue;
      }
      const cell = byPosition.get(key);
      const attrs = cell ? [`style="${htmlAttr(cellStyle(cell, styles))}"`].filter((attr) => attr !== 'style=""') : [];
      const merge = byMergeStart.get(key);
      if (merge?.rowspan && merge.rowspan > 1) {
        attrs.unshift(`rowspan="${merge.rowspan}"`);
      }
      if (merge?.colspan && merge.colspan > 1) {
        attrs.unshift(`colspan="${merge.colspan}"`);
      }
      const attrsText = attrs.length ? ` ${attrs.join(" ")}` : "";
      lines.push(`    <td${attrsText}>${renderCellValue(cell, styles)}</td>`);
    }
    lines.push("  </tr>");
  }
  lines.push("</table>");
  return lines.join("\n");
}

function tableBlocksForSheet(sheet: Dict): TableBlock[] {
  const cells: Dict[] = sheet.cells || [];
  const styles = styleMap(sheet);
  const meaningfulCells = cells.filter((cell) => isMeaningfulCell(cell, styles));
  const meaningfulByPosition = new Map(meaningfulCells.map((cell) => [`${cell.row},${cell.col}`, cell]));
  const merges = mergeRanges(sheet);
  const visited = new Set<string>();
  const blocks: TableBlock[] = [];

  for (const cell of meaningfulCells) {
    const startKey = `${cell.row},${cell.col}`;
    if (visited.has(startKey)) {
      continue;
    }
    const stack = [startKey];
    const component: Dict[] = [];
    visited.add(startKey);

    while (stack.length) {
      const current = stack.pop()!;
      const currentCell = meaningfulByPosition.get(current);
      if (!currentCell) {
        continue;
      }
      component.push(currentCell);
      const [row, col] = current.split(",").map(Number);
      for (const [nextRow, nextCol] of [[row - 1, col], [row + 1, col], [row, col - 1], [row, col + 1]]) {
        const nextKey = `${nextRow},${nextCol}`;
        if (meaningfulByPosition.has(nextKey) && !visited.has(nextKey)) {
          visited.add(nextKey);
          stack.push(nextKey);
        }
      }
    }

    const block = tableBlockFromComponent(component, cells, merges);
    if (block) {
      blocks.push(block);
    }
  }

  return blocks.sort((left, right) => left.bounds.minRow - right.bounds.minRow || left.bounds.minCol - right.bounds.minCol);
}

function tableBlockFromComponent(component: Dict[], allCells: Dict[], merges: MergeRange[]): TableBlock | undefined {
  const valuedCells = component.filter((cell) => cellText(cell).trim());
  if (!valuedCells.length) {
    return undefined;
  }
  const bounds: TableBounds = {
    minRow: Math.min(...valuedCells.map((cell) => Number(cell.row))),
    maxRow: Math.max(...valuedCells.map((cell) => Number(cell.row))),
    minCol: Math.min(...valuedCells.map((cell) => Number(cell.col))),
    maxCol: Math.max(...valuedCells.map((cell) => Number(cell.col)))
  };
  const blockMerges = merges.filter((merge) => (
    merge.startRow >= bounds.minRow &&
    merge.startRow <= bounds.maxRow &&
    merge.startCol >= bounds.minCol &&
    merge.startCol <= bounds.maxCol
  ));
  for (const merge of blockMerges) {
    bounds.minRow = Math.min(bounds.minRow, merge.startRow);
    bounds.maxRow = Math.max(bounds.maxRow, merge.endRow);
    bounds.minCol = Math.min(bounds.minCol, merge.startCol);
    bounds.maxCol = Math.max(bounds.maxCol, merge.endCol);
  }
  const byPosition = new Map(allCells.map((cell) => [`${cell.row},${cell.col}`, cell]));
  const cells: Dict[] = [];
  for (let row = bounds.minRow; row <= bounds.maxRow; row += 1) {
    for (let col = bounds.minCol; col <= bounds.maxCol; col += 1) {
      const cell = byPosition.get(`${row},${col}`);
      if (cell) {
        cells.push(cell);
      }
    }
  }
  return { bounds, cells, merges: blockMerges };
}

function isMeaningfulCell(cell: Dict, styles: Record<number, Dict>): boolean {
  if (cellText(cell).trim()) {
    return true;
  }
  const style = styles[Number(cell.style_id || 0)] || {};
  return Boolean(style.fill_color || style.border_color || style.apply_border);
}

function styleMap(sheet: Dict): Record<number, Dict> {
  return Object.fromEntries((sheet.styles || []).map((style: Dict) => [Number(style.style_id), style]));
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
    border: style.apply_border || style.border_color ? `1px solid ${style.border_color || "#D9D9D9"}` : undefined,
    background: style.fill_color || undefined,
    "text-align": style.horizontal_alignment ? cssHorizontalAlign(String(style.horizontal_alignment || "")) : undefined,
    "vertical-align": style.vertical_alignment ? cssVerticalAlign(String(style.vertical_alignment || "")) : undefined,
    "font-weight": font.bold ? "700" : undefined,
    color: font.color,
    "white-space": style.wrap_text ? "pre-wrap" : undefined
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
  return images.map((image: Dict, index: number) => {
    const sourcePath = String(image.path);
    const ext = imageExtension(sourcePath);
    const fileName = `${sheetFileName}-${index + 1}${ext}`;
    const destination = path.join(ctx.resourceDir, fileName);
    fs.writeFileSync(destination, readBuffer(ctx.zip, sourcePath));
    const relPath = path.relative(ctx.linkBaseDir, destination).split(path.sep).join("/");
    const alt = htmlAttr(String(image.name || path.parse(fileName).name));
    const width = Number(image.width || 0);
    const displayWidth = width > 0 ? Math.min(Math.max(Math.round(width), 320), 640) : 640;
    const html = `<img src="${htmlAttr(relPath)}" alt="${alt}" width="${displayWidth}">`;
    return {
      x: Number(image.x || 0),
      y: Number(image.y || 0),
      html
    };
  });
}

function mermaidSequenceDiagram(sheet: Dict): MermaidBlock | undefined {
  const vertices = (sheet.drawings || [])
    .filter((drawing: Dict) => drawing.kind === "vertex" && String(drawing.text || "").trim())
    .sort((left: Dict, right: Dict) => Number(left.x || 0) - Number(right.x || 0));
  const edges = (sheet.drawings || [])
    .filter((drawing: Dict) => drawing.kind === "edge" && drawing.sourceId && drawing.targetId)
    .sort((left: Dict, right: Dict) => Number(left.y || 0) - Number(right.y || 0));
  if (vertices.length < 2 || !edges.length) {
    return undefined;
  }

  const diagramItems = [...vertices, ...edges];
  const x = Math.min(...diagramItems.map((drawing: Dict) => Number(drawing.x || 0)));
  const y = Math.min(...diagramItems.map((drawing: Dict) => Number(drawing.y || 0)));
  const aliases = new Map<string, string>();
  const lines = ["sequenceDiagram"];
  vertices.forEach((vertex: Dict, index: number) => {
    const alias = `P${index + 1}`;
    aliases.set(String(vertex.id), alias);
    lines.push(`  participant ${alias} as ${mermaidText(String(vertex.text || vertex.name || alias))}`);
  });

  let messageCount = 0;
  for (const edge of edges) {
    const source = aliases.get(String(edge.sourceId));
    const target = aliases.get(String(edge.targetId));
    if (!source || !target || source === target) {
      continue;
    }
    const message = mermaidText(String(edge.text || edge.name || "message"));
    lines.push(`  ${source}->>${target}: ${message}`);
    messageCount += 1;
  }
  return messageCount ? { x, y, code: lines.join("\n") } : undefined;
}

function mermaidText(value: string): string {
  return value.replace(/\r?\n/g, " ").replace(/:/g, "：").trim() || "message";
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
