import * as fs from "fs";
import * as path from "path";
import AdmZip = require("adm-zip");
import {
  asArray,
  entries,
  findAll,
  first,
  hasEntry,
  quoteAttr,
  readBuffer,
  readXml,
  resolvePackagePath,
  round,
  textContent,
  xmlEscape
} from "./shared";

const EMU_PER_INCH = 914400;
const PX_PER_INCH = 96;
const DEFAULT_COL_WIDTH = 12.63;
const DEFAULT_ROW_HEIGHT_PT = 15.75;
const SHEET_ORIGIN_X = 40;
const SHEET_ORIGIN_Y = 40;

type Dict = Record<string, any>;

export async function convertXlsxToDrawio(sourcePath: string): Promise<string> {
  const drawioPath = path.join(path.dirname(sourcePath), `${path.parse(sourcePath).name}.drawio`);
  const intermediate = buildIntermediate(sourcePath);
  fs.writeFileSync(drawioPath, makeDrawio(intermediate), "utf8");
  return drawioPath;
}

function emuToPx(value: unknown): number {
  return round((Number(value) / EMU_PER_INCH) * PX_PER_INCH);
}

function colNameToIndex(colName: string): number {
  let index = 0;
  for (const ch of colName.toUpperCase()) {
    index = index * 26 + ch.charCodeAt(0) - "A".charCodeAt(0) + 1;
  }
  return index;
}

function splitCellRef(ref: string): [number, number] {
  const match = ref.match(/^([A-Z]+)([0-9]+)/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  return [colNameToIndex(match[1]), Number(match[2])];
}

function readRelationships(zip: AdmZip, name: string): Record<string, Dict> {
  if (!hasEntry(zip, name)) {
    return {};
  }
  const root = readXml(zip, name);
  const relationships: Record<string, Dict> = {};
  for (const rel of asArray<any>(root.Relationships?.Relationship)) {
    relationships[String(rel.Id)] = {
      type: String(rel.Type || ""),
      target: String(rel.Target || "")
    };
  }
  return relationships;
}

function readSharedStrings(zip: AdmZip): string[] {
  if (!hasEntry(zip, "xl/sharedStrings.xml")) {
    return [];
  }
  const root = readXml(zip, "xl/sharedStrings.xml");
  return asArray<any>(root.sst?.si).map((si) => findAll(si, "t").map(textContent).join(""));
}

function parseWorkbook(zip: AdmZip): Dict[] {
  const workbook = readXml(zip, "xl/workbook.xml");
  const workbookRels = readRelationships(zip, "xl/_rels/workbook.xml.rels");
  return asArray<any>(workbook.workbook?.sheets?.sheet).map((sheet) => {
    const relId = String(sheet.id || "");
    const target = workbookRels[relId]?.target || "";
    return {
      name: String(sheet.name || ""),
      sheet_id: String(sheet.sheetId || ""),
      rel_id: relId,
      path: target ? resolvePackagePath("xl/workbook.xml", target) : ""
    };
  });
}

function displayValueFromRaw(cellType: string, rawValue: string, sharedStrings: string[]): string {
  if (cellType === "s" && rawValue) {
    return sharedStrings[Number(rawValue)] || "";
  }
  if (cellType === "inlineStr") {
    return rawValue;
  }
  if (cellType === "n" && rawValue) {
    const number = Number(rawValue);
    if (Number.isFinite(number) && Number.isInteger(number)) {
      return String(number);
    }
  }
  return rawValue;
}

function normalizeRgb(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }
  const rgb = value.slice(-6);
  return rgb.length === 6 ? `#${rgb}` : undefined;
}

function parseStyles(zip: AdmZip): Dict[] {
  if (!hasEntry(zip, "xl/styles.xml")) {
    return [];
  }
  const root = readXml(zip, "xl/styles.xml").styleSheet;
  const fills = asArray<any>(root?.fills?.fill).map((fill) => ({
    color: normalizeRgb(first<any>(fill.patternFill, "fgColor")?.rgb || fill.patternFill?.fgColor?.rgb)
  }));
  const borders = asArray<any>(root?.borders?.border).map((border) => {
    let borderColor: string | undefined;
    for (const sideName of ["left", "right", "top", "bottom"]) {
      const color = border?.[sideName]?.color;
      borderColor = normalizeRgb(color?.rgb) || borderColor;
    }
    return { color: borderColor };
  });
  return asArray<any>(root?.cellXfs?.xf).map((xf, index) => {
    const fillId = Number(xf.fillId || 0);
    const borderId = Number(xf.borderId || 0);
    return {
      style_id: index,
      font_id: Number(xf.fontId || 0),
      fill_id: fillId,
      fill_color: fills[fillId]?.color,
      border_id: borderId,
      border_color: borders[borderId]?.color,
      num_fmt_id: Number(xf.numFmtId || 0),
      apply_border: String(xf.applyBorder || "") === "1",
      apply_alignment: String(xf.applyAlignment || "") === "1"
    };
  });
}

function parseCells(sheetRoot: Dict, sharedStrings: string[]): [Dict[], Dict[], Dict[]] {
  const rows: Dict[] = [];
  const cells: Dict[] = [];
  const cols = asArray<any>(sheetRoot.worksheet?.cols?.col).map((col) => ({
    min: Number(col.min || 0),
    max: Number(col.max || 0),
    width: Number(col.width || 0),
    custom_width: String(col.customWidth || "") === "1"
  }));

  for (const row of asArray<any>(sheetRoot.worksheet?.sheetData?.row)) {
    const rowIndex = Number(row.r);
    rows.push({
      index: rowIndex,
      height: row.ht !== undefined ? Number(row.ht) : undefined
    });
    for (const cell of asArray<any>(row.c)) {
      const ref = String(cell.r);
      const [colIndex, rowNumber] = splitCellRef(ref);
      const cellType = String(cell.t || "n");
      const rawValue = cellType === "inlineStr" ? findAll(cell, "t").map(textContent).join("") : String(textContent(cell.v) || "");
      cells.push({
        ref,
        row: rowNumber,
        col: colIndex,
        style_id: Number(cell.s || 0),
        type: cellType,
        raw_value: rawValue,
        value: displayValueFromRaw(cellType, rawValue, sharedStrings)
      });
    }
  }
  return [rows, cols, cells];
}

function parseMergedCells(sheetRoot: Dict): Dict[] {
  return asArray<any>(sheetRoot.worksheet?.mergeCells?.mergeCell)
    .map((mergeCell) => String(mergeCell.ref || ""))
    .filter((ref) => ref.includes(":"))
    .map((ref) => {
      const [start, end] = ref.split(":", 2);
      return { ref, start, end };
    });
}

function textFromDrawing(node: Dict): string {
  return findAll(node, "t").map(textContent).join("");
}

function findColor(node: Dict, colorParent: string): string | undefined {
  const parents = findAll(node, colorParent);
  for (const parent of parents) {
    const color = first<any>(parent, "srgbClr") || parent.srgbClr;
    if (color?.val) {
      return `#${color.val}`;
    }
  }
  return undefined;
}

function imageDataUri(zip: AdmZip, imagePath: string): string {
  const suffix = path.extname(imagePath).toLowerCase().replace(/^\./, "");
  const mime = suffix === "jpg" || suffix === "jpeg" ? "jpeg" : suffix || "octet-stream";
  return `data:image/${mime},${readBuffer(zip, imagePath).toString("base64")}`;
}

function excelColWidthToPx(width: number): number {
  return round(width * 7 + 5);
}

function excelRowHeightToPx(heightPt: number): number {
  return round((heightPt * PX_PER_INCH) / 72);
}

function colWidthMapFromCols(cols: Dict[]): Record<number, number> {
  const widths: Record<number, number> = {};
  for (const col of cols) {
    for (let index = Number(col.min); index <= Number(col.max); index += 1) {
      widths[index] = excelColWidthToPx(Number(col.width || DEFAULT_COL_WIDTH));
    }
  }
  return widths;
}

function rowHeightMapFromRows(rows: Dict[]): Record<number, number> {
  const heights: Record<number, number> = {};
  for (const row of rows) {
    heights[Number(row.index)] = excelRowHeightToPx(Number(row.height || DEFAULT_ROW_HEIGHT_PT));
  }
  return heights;
}

function colWidthMap(sheet: Dict): Record<number, number> {
  return colWidthMapFromCols(sheet.cols || []);
}

function rowHeightMap(sheet: Dict): Record<number, number> {
  return rowHeightMapFromRows(sheet.rows || []);
}

function sheetXFromZeroBasedCol(col: number, colWidths: Record<number, number>): number {
  const defaultColPx = excelColWidthToPx(DEFAULT_COL_WIDTH);
  let sum = 0;
  for (let i = 1; i <= col; i += 1) {
    sum += colWidths[i] || defaultColPx;
  }
  return SHEET_ORIGIN_X + sum;
}

function sheetYFromZeroBasedRow(row: number, rowHeights: Record<number, number>): number {
  const defaultRowPx = excelRowHeightToPx(DEFAULT_ROW_HEIGHT_PT);
  let sum = 0;
  for (let i = 1; i <= row; i += 1) {
    sum += rowHeights[i] || defaultRowPx;
  }
  return SHEET_ORIGIN_Y + sum;
}

function parseAnchor(anchor: Dict, colWidths: Record<number, number>, rowHeights: Record<number, number>, type: string): Dict {
  const fromNode = anchor.from;
  const ext = anchor.ext;
  const result: Dict = { type, from: undefined, ext: undefined };
  if (fromNode) {
    const from = {
      col: Number(fromNode.col || 0),
      row: Number(fromNode.row || 0),
      colOff_emu: Number(fromNode.colOff || 0),
      rowOff_emu: Number(fromNode.rowOff || 0)
    };
    result.from = from;
    result.x = round(sheetXFromZeroBasedCol(from.col, colWidths) + emuToPx(from.colOff_emu));
    result.y = round(sheetYFromZeroBasedRow(from.row, rowHeights) + emuToPx(from.rowOff_emu));
  }
  if (ext) {
    result.ext = {
      cx_emu: Number(ext.cx || 0),
      cy_emu: Number(ext.cy || 0),
      width: emuToPx(ext.cx || 0),
      height: emuToPx(ext.cy || 0)
    };
  }
  return result;
}

function parseGroupTransform(group: Dict): Dict {
  const xfrm = group.grpSpPr?.xfrm;
  return {
    off_x: Number(xfrm?.off?.x || 0),
    off_y: Number(xfrm?.off?.y || 0),
    ch_off_x: Number(xfrm?.chOff?.x || 0),
    ch_off_y: Number(xfrm?.chOff?.y || 0)
  };
}

function parseDrawingShape(node: Dict, anchorInfo: Dict, kind: "vertex" | "edge", groupTransform?: Dict): Dict {
  const nv = findAll(node, "cNvPr")[0];
  const xfrm = findAll(node, "xfrm")[0];
  const geom = findAll(node, "prstGeom")[0];
  const tail = findAll(node, "tailEnd")[0];
  const head = findAll(node, "headEnd")[0];
  const startCxn = findAll(node, "stCxn")[0];
  const endCxn = findAll(node, "endCxn")[0];
  const shape: Dict = {
    id: nv?.id ? String(nv.id) : "",
    name: nv?.name ? String(nv.name) : "",
    kind,
    preset: geom?.prst ? String(geom.prst) : "",
    text: textFromDrawing(node),
    fill: findColor(node, "solidFill"),
    stroke: findColor(node, "ln"),
    headEnd: head?.type,
    tailEnd: tail?.type,
    startConnectionId: startCxn?.id,
    endConnectionId: endCxn?.id
  };
  const off = xfrm?.off;
  const ext = xfrm?.ext;
  if (off) {
    const rawXEmu = Number(off.x || 0);
    const rawYEmu = Number(off.y || 0);
    const localXEmu = rawXEmu - Number(groupTransform?.ch_off_x || 0);
    const localYEmu = rawYEmu - Number(groupTransform?.ch_off_y || 0);
    shape.raw_x_emu = rawXEmu;
    shape.raw_y_emu = rawYEmu;
    shape.local_x_emu = localXEmu;
    shape.local_y_emu = localYEmu;
    shape.x_emu = localXEmu;
    shape.y_emu = localYEmu;
    shape.x = round(Number(anchorInfo.x || 0) + emuToPx(localXEmu));
    shape.y = round(Number(anchorInfo.y || 0) + emuToPx(localYEmu));
  }
  if (ext) {
    shape.width_emu = Number(ext.cx || 0);
    shape.height_emu = Number(ext.cy || 0);
    shape.width = emuToPx(ext.cx || 0);
    shape.height = emuToPx(ext.cy || 0);
  }
  return shape;
}

function parseDrawingPicture(zip: AdmZip, node: Dict, anchorInfo: Dict, drawingRels: Record<string, Dict>, drawingPath: string, pictureIndex: number): Dict {
  const nv = findAll(node, "cNvPr")[0];
  const blip = findAll(node, "blip")[0];
  const relId = blip?.embed ? String(blip.embed) : "";
  const target = drawingRels[relId]?.target || "";
  const imagePath = target ? resolvePackagePath(drawingPath, target) : "";
  const ext = anchorInfo.ext || {};
  return {
    id: `pic-${pictureIndex}`,
    name: nv?.name ? String(nv.name) : "",
    kind: "image",
    rel_id: relId,
    path: imagePath,
    x: anchorInfo.x || 0,
    y: anchorInfo.y || 0,
    width: ext.width || 120,
    height: ext.height || 80,
    data_uri: imagePath ? imageDataUri(zip, imagePath) : ""
  };
}

function parseDrawings(zip: AdmZip, drawingPath: string, rows: Dict[], cols: Dict[]): [Dict[], Dict[]] {
  if (!hasEntry(zip, drawingPath)) {
    return [[], []];
  }
  const root = readXml(zip, drawingPath).wsDr;
  const relPath = `${path.posix.dirname(drawingPath)}/_rels/${path.posix.basename(drawingPath)}.rels`;
  const drawingRels = readRelationships(zip, relPath);
  const colWidths = colWidthMapFromCols(cols);
  const rowHeights = rowHeightMapFromRows(rows);
  const anchors: Dict[] = [];
  const drawings: Dict[] = [];
  let pictureIndex = 1;

  for (const anchorType of ["oneCellAnchor", "twoCellAnchor", "absoluteAnchor"]) {
    for (const anchor of asArray<any>(root?.[anchorType])) {
      const anchorInfo = parseAnchor(anchor, colWidths, rowHeights, anchorType);
      anchors.push(anchorInfo);
      for (const child of asArray<any>(anchor.sp)) {
        drawings.push(parseDrawingShape(child, anchorInfo, "vertex"));
      }
      for (const child of asArray<any>(anchor.cxnSp)) {
        drawings.push(parseDrawingShape(child, anchorInfo, "edge"));
      }
      for (const child of asArray<any>(anchor.pic)) {
        drawings.push(parseDrawingPicture(zip, child, anchorInfo, drawingRels, drawingPath, pictureIndex));
        pictureIndex += 1;
      }
      for (const group of asArray<any>(anchor.grpSp)) {
        const groupTransform = parseGroupTransform(group);
        for (const child of asArray<any>(group.sp)) {
          drawings.push(parseDrawingShape(child, anchorInfo, "vertex", groupTransform));
        }
        for (const child of asArray<any>(group.cxnSp)) {
          drawings.push(parseDrawingShape(child, anchorInfo, "edge", groupTransform));
        }
        for (const child of asArray<any>(group.pic)) {
          drawings.push(parseDrawingPicture(zip, child, anchorInfo, drawingRels, drawingPath, pictureIndex));
          pictureIndex += 1;
        }
      }
    }
  }
  inferEdgeSources(drawings);
  return [drawings, anchors];
}

function inferEdgeSources(drawings: Dict[]): void {
  const vertices = drawings.filter((drawing) => drawing.kind === "vertex");
  for (const edge of drawings.filter((drawing) => drawing.kind === "edge")) {
    edge.sourceId = edge.startConnectionId || nearestVertexId(edge, vertices, false);
    edge.targetId = edge.endConnectionId || nearestVertexId(edge, vertices, true);
  }
}

function nearestVertexId(edge: Dict, vertices: Dict[], useEnd: boolean): string | undefined {
  let ex = Number(edge.x || 0);
  let ey = Number(edge.y || 0);
  if (useEnd) {
    ex += Number(edge.width || 0);
    ey += Number(edge.height || 0);
  }
  let bestId: string | undefined;
  let bestDistance: number | undefined;
  for (const vertex of vertices) {
    const vx = Number(vertex.x || 0) + Number(vertex.width || 0) / 2;
    const vy = Number(vertex.y || 0) + Number(vertex.height || 0) / 2;
    const distance = (vx - ex) ** 2 + (vy - ey) ** 2;
    if (bestDistance === undefined || distance < bestDistance) {
      bestDistance = distance;
      bestId = String(vertex.id);
    }
  }
  return bestId;
}

function readImages(zip: AdmZip): Dict[] {
  return entries(zip)
    .filter((name) => name.startsWith("xl/media/"))
    .map((name) => {
      const suffix = path.extname(name).toLowerCase().replace(/^\./, "");
      const mime = suffix === "jpg" || suffix === "jpeg" ? "jpeg" : suffix || "octet-stream";
      return {
        path: name,
        mime: `image/${mime}`,
        base64: readBuffer(zip, name).toString("base64")
      };
    });
}

function buildIntermediate(xlsxPath: string): Dict {
  const zip = new AdmZip(xlsxPath);
  const sharedStrings = readSharedStrings(zip);
  const styles = parseStyles(zip);
  const sheetDefs = parseWorkbook(zip);
  const images = readImages(zip);
  const sheets: Dict[] = [];
  for (const sheetDef of sheetDefs) {
    const sheetRoot = readXml(zip, sheetDef.path);
    const [rows, cols, cells] = parseCells(sheetRoot, sharedStrings);
    const relPath = `${path.posix.dirname(sheetDef.path)}/_rels/${path.posix.basename(sheetDef.path)}.rels`;
    const sheetRels = readRelationships(zip, relPath);
    const drawings: Dict[] = [];
    const anchors: Dict[] = [];
    for (const rel of Object.values(sheetRels)) {
      if (String(rel.type).endsWith("/drawing")) {
        const drawingPath = resolvePackagePath(sheetDef.path, String(rel.target));
        const [drawingItems, anchorItems] = parseDrawings(zip, drawingPath, rows, cols);
        drawings.push(...drawingItems);
        anchors.push(...anchorItems);
      }
    }
    sheets.push({
      name: sheetDef.name,
      sheet_id: sheetDef.sheet_id,
      path: sheetDef.path,
      rows,
      cols,
      cells,
      styles,
      merged_cells: parseMergedCells(sheetRoot),
      drawings,
      images,
      anchors
    });
  }
  return {
    workbook: {
      source: xlsxPath,
      format: "xlsx-zip",
      sheets_count: sheets.length
    },
    sheets
  };
}

function drawioShapeStyle(shape: Dict): string {
  const fill = shape.fill || "#FFFFFF";
  const stroke = shape.stroke || "#000000";
  const rounded = shape.preset === "flowChartAlternateProcess" ? "1;arcSize=12" : "0";
  return `rounded=${rounded};whiteSpace=wrap;html=1;fillColor=${fill};strokeColor=${stroke};`;
}

function cellGeometry(cell: Dict, colWidths: Record<number, number>, rowHeights: Record<number, number>): Dict {
  const col = Number(cell.col);
  const row = Number(cell.row);
  const defaultColPx = excelColWidthToPx(DEFAULT_COL_WIDTH);
  const defaultRowPx = excelRowHeightToPx(DEFAULT_ROW_HEIGHT_PT);
  let x = SHEET_ORIGIN_X;
  for (let i = 1; i < col; i += 1) {
    x += colWidths[i] || defaultColPx;
  }
  let y = SHEET_ORIGIN_Y;
  for (let i = 1; i < row; i += 1) {
    y += rowHeights[i] || defaultRowPx;
  }
  return {
    x: round(x),
    y: round(y),
    width: colWidths[col] || defaultColPx,
    height: rowHeights[row] || defaultRowPx
  };
}

function isTableCell(cell: Dict, styles: Record<number, Dict>): boolean {
  const style = styles[Number(cell.style_id || 0)] || {};
  return Boolean(style.apply_border || style.fill_color);
}

function drawioCellStyle(cell: Dict, styles: Record<number, Dict>): string {
  const style = styles[Number(cell.style_id || 0)] || {};
  if (isTableCell(cell, styles)) {
    const fill = style.fill_color || "#FFFFFF";
    const stroke = style.border_color || "#000000";
    return `rounded=0;whiteSpace=wrap;html=1;fillColor=${fill};strokeColor=${stroke};align=center;verticalAlign=middle;fontSize=11;`;
  }
  return "text;html=1;fillColor=none;strokeColor=none;align=center;verticalAlign=middle;fontSize=11;";
}

function findTableComponents(cells: Dict[], styles: Record<number, Dict>): Dict[][] {
  const tableCells = new Map<string, Dict>();
  for (const cell of cells) {
    if (isTableCell(cell, styles)) {
      tableCells.set(`${cell.row},${cell.col}`, cell);
    }
  }
  const visited = new Set<string>();
  const components: Dict[][] = [];
  for (const key of tableCells.keys()) {
    if (visited.has(key)) {
      continue;
    }
    const stack = [key];
    visited.add(key);
    const component: Dict[] = [];
    while (stack.length) {
      const current = stack.pop()!;
      component.push(tableCells.get(current)!);
      const [row, col] = current.split(",").map(Number);
      for (const neighbor of [[row - 1, col], [row + 1, col], [row, col - 1], [row, col + 1]]) {
        const neighborKey = `${neighbor[0]},${neighbor[1]}`;
        if (tableCells.has(neighborKey) && !visited.has(neighborKey)) {
          visited.add(neighborKey);
          stack.push(neighborKey);
        }
      }
    }
    components.push(component);
  }
  return components;
}

function tableGeometry(component: Dict[], colWidths: Record<number, number>, rowHeights: Record<number, number>): Dict {
  const rows = component.map((cell) => Number(cell.row));
  const cols = component.map((cell) => Number(cell.col));
  const minRow = Math.min(...rows);
  const maxRow = Math.max(...rows);
  const minCol = Math.min(...cols);
  const maxCol = Math.max(...cols);
  const topLeft = cellGeometry({ row: minRow, col: minCol }, colWidths, rowHeights);
  let width = 0;
  for (let col = minCol; col <= maxCol; col += 1) {
    width += colWidths[col] || excelColWidthToPx(DEFAULT_COL_WIDTH);
  }
  let height = 0;
  for (let row = minRow; row <= maxRow; row += 1) {
    height += rowHeights[row] || excelRowHeightToPx(DEFAULT_ROW_HEIGHT_PT);
  }
  return { ...topLeft, width: round(width), height: round(height), min_row: minRow, max_row: maxRow, min_col: minCol, max_col: maxCol };
}

function htmlTdStyle(cell: Dict | undefined, styles: Record<number, Dict>, width: number, height: number): string {
  const style = cell ? styles[Number(cell.style_id || 0)] || {} : {};
  const fill = style.fill_color || "#FFFFFF";
  const stroke = style.border_color || "#000000";
  return `border:1px solid ${stroke};background:${fill};text-align:center;vertical-align:middle;width:${width}px;height:${height}px;font-size:11px;padding:0;box-sizing:border-box;`;
}

function buildHtmlTable(component: Dict[], styles: Record<number, Dict>, colWidths: Record<number, number>, rowHeights: Record<number, number>): string {
  const geometry = tableGeometry(component, colWidths, rowHeights);
  const byPosition = new Map(component.map((cell) => [`${cell.row},${cell.col}`, cell]));
  const lines = ['<table style="border-collapse:collapse;table-layout:fixed;width:100%;height:100%;font-family:Arial,sans-serif;">'];
  for (let row = geometry.min_row; row <= geometry.max_row; row += 1) {
    const rowHeight = rowHeights[row] || excelRowHeightToPx(DEFAULT_ROW_HEIGHT_PT);
    lines.push("<tr>");
    for (let col = geometry.min_col; col <= geometry.max_col; col += 1) {
      const colWidth = colWidths[col] || excelColWidthToPx(DEFAULT_COL_WIDTH);
      const cell = byPosition.get(`${row},${col}`);
      const value = cell ? xmlEscape(String(cell.value || "")) : "";
      lines.push(`<td style="${htmlTdStyle(cell, styles, colWidth, rowHeight)}">${value}</td>`);
    }
    lines.push("</tr>");
  }
  lines.push("</table>");
  return lines.join("");
}

function appendHtmlTableVertices(xmlCells: string[], sheet: Dict, styles: Record<number, Dict>, colWidths: Record<number, number>, rowHeights: Record<number, number>): Set<string> {
  const tableCellRefs = new Set<string>();
  for (const component of findTableComponents(sheet.cells || [], styles)) {
    if (!component.length) {
      continue;
    }
    const geometry = tableGeometry(component, colWidths, rowHeights);
    const html = buildHtmlTable(component, styles, colWidths, rowHeights);
    const tableId = `table-${geometry.min_row}-${geometry.min_col}-${geometry.max_row}-${geometry.max_col}`;
    xmlCells.push(`        <mxCell id="${tableId}" value=${quoteAttr(html)} style=${quoteAttr("html=1;whiteSpace=wrap;overflow=fill;rounded=0;fillColor=none;strokeColor=none;")} vertex="1" parent="1">`);
    xmlCells.push(`          <mxGeometry x="${geometry.x}" y="${geometry.y}" width="${geometry.width}" height="${geometry.height}" as="geometry"/>`);
    xmlCells.push("        </mxCell>");
    for (const cell of component) {
      tableCellRefs.add(String(cell.ref));
    }
  }
  return tableCellRefs;
}

function appendCellVertices(xmlCells: string[], sheet: Dict): void {
  const styles = Object.fromEntries((sheet.styles || []).map((style: Dict) => [Number(style.style_id), style]));
  const colWidths = colWidthMap(sheet);
  const rowHeights = rowHeightMap(sheet);
  const tableCellRefs = appendHtmlTableVertices(xmlCells, sheet, styles, colWidths, rowHeights);
  for (const cell of sheet.cells || []) {
    if (tableCellRefs.has(String(cell.ref))) {
      continue;
    }
    const geometry = cellGeometry(cell, colWidths, rowHeights);
    xmlCells.push(`        <mxCell id="cell-${cell.ref}" value=${quoteAttr(String(cell.value || ""))} style=${quoteAttr(drawioCellStyle(cell, styles))} vertex="1" parent="1">`);
    xmlCells.push(`          <mxGeometry x="${geometry.x}" y="${geometry.y}" width="${geometry.width}" height="${geometry.height}" as="geometry"/>`);
    xmlCells.push("        </mxCell>");
  }
}

function appendDrawingVertices(xmlCells: string[], sheet: Dict): void {
  for (const shape of sheet.drawings || []) {
    if (shape.kind !== "vertex") {
      continue;
    }
    xmlCells.push(`        <mxCell id="shape-${shape.id}" value=${quoteAttr(String(shape.text || ""))} style=${quoteAttr(drawioShapeStyle(shape))} vertex="1" parent="1">`);
    xmlCells.push(`          <mxGeometry x="${shape.x || 0}" y="${shape.y || 0}" width="${shape.width || 120}" height="${shape.height || 60}" as="geometry"/>`);
    xmlCells.push("        </mxCell>");
  }
}

function appendDrawingEdges(xmlCells: string[], sheet: Dict): void {
  for (const shape of sheet.drawings || []) {
    if (shape.kind !== "edge") {
      continue;
    }
    const source = shape.sourceId ? `shape-${shape.sourceId}` : "";
    const target = shape.targetId ? `shape-${shape.targetId}` : "";
    const sourceAttr = source ? ` source="${source}"` : "";
    const targetAttr = target ? ` target="${target}"` : "";
    const stroke = shape.stroke || "#000000";
    xmlCells.push(`        <mxCell id="edge-${shape.id}" value="" style=${quoteAttr(`edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;strokeColor=${stroke};endArrow=classic;`)} edge="1" parent="1"${sourceAttr}${targetAttr}>`);
    xmlCells.push('          <mxGeometry relative="1" as="geometry"/>');
    xmlCells.push("        </mxCell>");
  }
}

function appendImageVertices(xmlCells: string[], sheet: Dict): void {
  for (const image of sheet.drawings || []) {
    if (image.kind !== "image") {
      continue;
    }
    xmlCells.push(`        <mxCell id="image-${image.id}" value="" style=${quoteAttr(`shape=image;html=1;imageAspect=0;aspect=fixed;image=${image.data_uri || ""};`)} vertex="1" parent="1">`);
    xmlCells.push(`          <mxGeometry x="${image.x || 0}" y="${image.y || 0}" width="${image.width || 120}" height="${image.height || 80}" as="geometry"/>`);
    xmlCells.push("        </mxCell>");
  }
}

function drawioDiagramXml(sheet: Dict, index: number): string[] {
  const cells = [
    `  <diagram id=${quoteAttr(`POC06-${index}`)} name=${quoteAttr(String(sheet.name || `Sheet${index}`))}>`,
    '    <mxGraphModel dx="1000" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169">',
    "      <root>",
    '        <mxCell id="0"/>',
    '        <mxCell id="1" parent="0"/>'
  ];
  appendCellVertices(cells, sheet);
  appendImageVertices(cells, sheet);
  appendDrawingVertices(cells, sheet);
  appendDrawingEdges(cells, sheet);
  cells.push("      </root>", "    </mxGraphModel>", "  </diagram>");
  return cells;
}

function makeDrawio(intermediate: Dict): string {
  const cells = ['<?xml version="1.0" encoding="UTF-8"?>', '<mxfile host="app.diagrams.net">'];
  (intermediate.sheets || []).forEach((sheet: Dict, index: number) => {
    for (const line of drawioDiagramXml(sheet, index + 1)) {
      cells.push(line);
    }
  });
  cells.push("</mxfile>");
  return cells.join("\n");
}
