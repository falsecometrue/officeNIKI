import * as fs from "fs";
import * as path from "path";
import AdmZip = require("adm-zip");
import { asArray, childEntries, findAll, first, hasEntry, readBuffer, readXml, textContent, XmlNode } from "./shared";

const EMU_PER_PIXEL = 9525;

type Relationship = {
  type: string;
  target: string;
};

type Context = {
  sourceDocx: string;
  outputDir: string;
  resourceDir: string;
  rels: Record<string, Relationship>;
  styles: Record<string, StyleInfo>;
  numbering: NumberingInfo;
  copiedResources: Record<string, string>;
};

type Drawing = Record<string, any>;
type Block = Record<string, any>;
type StyleInfo = {
  run?: Record<string, any>;
  paragraph?: Record<string, any>;
};
type NumberingLevel = {
  format: string;
  text: string;
};
type NumberingInfo = {
  nums: Record<string, string>;
  levels: Record<string, Record<string, NumberingLevel>>;
};

export async function convertDocxToMarkdown(sourcePath: string): Promise<string> {
  const sourceName = path.parse(sourcePath).name;
  const outputDir = path.join(path.dirname(sourcePath), sourceName);
  fs.mkdirSync(outputDir, { recursive: true });
  const markdownPath = path.join(outputDir, `${sourceName}.md`);
  const intermediate = parseDocx(sourcePath, outputDir);
  fs.writeFileSync(markdownPath, buildMarkdown(intermediate), "utf8");
  return markdownPath;
}

function parseRelationships(zip: AdmZip): Record<string, Relationship> {
  const root = readXml(zip, "word/_rels/document.xml.rels");
  const rels: Record<string, Relationship> = {};
  for (const rel of asArray<any>(root.Relationships?.Relationship)) {
    let target = String(rel.Target || "");
    if (target && !target.startsWith("/")) {
      target = `word/${target}`;
    } else if (target.startsWith("/")) {
      target = target.slice(1);
    }
    rels[String(rel.Id)] = {
      type: String(rel.Type || ""),
      target
    };
  }
  return rels;
}

function parseImageSize(data: Buffer, suffix: string): Record<string, number> | undefined {
  const lower = suffix.toLowerCase();
  if (lower === ".png" && data.length >= 24 && data.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    return { width_px: data.readUInt32BE(16), height_px: data.readUInt32BE(20) };
  }
  if ((lower === ".jpg" || lower === ".jpeg") && data.length > 10 && data[0] === 0xff && data[1] === 0xd8) {
    let i = 2;
    while (i < data.length - 9) {
      if (data[i] !== 0xff) {
        i += 1;
        continue;
      }
      const marker = data[i + 1];
      const length = data.readUInt16BE(i + 2);
      if (marker >= 0xc0 && marker <= 0xc3) {
        return { height_px: data.readUInt16BE(i + 5), width_px: data.readUInt16BE(i + 7) };
      }
      i += 2 + length;
    }
  }
  return undefined;
}

function copyResource(ctx: Context, zip: AdmZip, target: string): Record<string, any> {
  let relPath = ctx.copiedResources[target];
  if (!relPath) {
    const data = readBuffer(zip, target);
    let dest = path.join(ctx.resourceDir, path.basename(target));
    if (fs.existsSync(dest)) {
      const parsed = path.parse(dest);
      let n = 2;
      while (fs.existsSync(path.join(parsed.dir, `${parsed.name}_${n}${parsed.ext}`))) {
        n += 1;
      }
      dest = path.join(parsed.dir, `${parsed.name}_${n}${parsed.ext}`);
    }
    fs.writeFileSync(dest, data);
    relPath = path.relative(ctx.outputDir, dest).split(path.sep).join("/");
    ctx.copiedResources[target] = relPath;
  }

  const data = readBuffer(zip, target);
  const suffix = path.extname(target);
  return {
    source: target,
    path: relPath,
    content_type_hint: suffix.toLowerCase().replace(/^\./, ""),
    ...(parseImageSize(data, suffix) || {})
  };
}

function paragraphStyle(paragraph: XmlNode): string {
  const pstyle = first<any>(first<any>(paragraph, "pPr"), "pStyle");
  return String(pstyle?.val || "Normal");
}

function parseStyles(zip: AdmZip): Record<string, StyleInfo> {
  if (!hasEntry(zip, "word/styles.xml")) {
    return {};
  }
  const root = readXml(zip, "word/styles.xml");
  const styles: Record<string, StyleInfo> = {};
  for (const style of asArray<any>(root.styles?.style)) {
    const id = String(style.styleId || style.styleID || "");
    if (!id) {
      continue;
    }
    styles[id] = {
      run: parseRunProperties(first<any>(style, "rPr") || {}),
      paragraph: parseParagraphProperties(first<any>(style, "pPr") || {})
    };
  }
  return styles;
}

function parseNumbering(zip: AdmZip): NumberingInfo {
  if (!hasEntry(zip, "word/numbering.xml")) {
    return { nums: {}, levels: {} };
  }
  const root = readXml(zip, "word/numbering.xml");
  const numbering = root.numbering || {};
  const levels: Record<string, Record<string, NumberingLevel>> = {};
  for (const abstractNum of asArray<any>(numbering.abstractNum)) {
    const abstractNumId = String(abstractNum.abstractNumId || "");
    if (!abstractNumId) {
      continue;
    }
    levels[abstractNumId] = {};
    for (const lvl of asArray<any>(abstractNum.lvl)) {
      const ilvl = String(lvl.ilvl || "0");
      levels[abstractNumId][ilvl] = {
        format: String(first<any>(lvl, "numFmt")?.val || ""),
        text: String(first<any>(lvl, "lvlText")?.val || "")
      };
    }
  }

  const nums: Record<string, string> = {};
  for (const num of asArray<any>(numbering.num)) {
    const numId = String(num.numId || "");
    const abstractNumId = String(first<any>(num, "abstractNumId")?.val || "");
    if (numId && abstractNumId) {
      nums[numId] = abstractNumId;
    }
  }
  return { nums, levels };
}

function parseParagraphProperties(ppr: XmlNode | undefined): Record<string, any> {
  const jc = first<any>(ppr, "jc");
  const align = String(jc?.val || "");
  return align ? { align } : {};
}

function paragraphFormat(paragraph: XmlNode, styles: Record<string, StyleInfo>): Record<string, any> {
  const ppr = first<any>(paragraph, "pPr");
  const style = paragraphStyle(paragraph);
  return {
    ...(styles[style]?.paragraph || {}),
    ...parseParagraphProperties(ppr)
  };
}

function paragraphNumbering(paragraph: XmlNode, numbering: NumberingInfo): Record<string, any> | undefined {
  const numPr = first<any>(first<any>(paragraph, "pPr"), "numPr");
  const numId = String(first<any>(numPr, "numId")?.val || "");
  const ilvl = String(first<any>(numPr, "ilvl")?.val || "0");
  const abstractNumId = numbering.nums[numId];
  const level = abstractNumId ? numbering.levels[abstractNumId]?.[ilvl] : undefined;
  if (!numId || !level) {
    return undefined;
  }
  return {
    numId,
    ilvl: Number(ilvl) || 0,
    format: level.format,
    text: level.text
  };
}

function parseRunProperties(rpr: XmlNode | undefined): Record<string, any> {
  const color = first<any>(rpr, "color");
  const size = first<any>(rpr, "sz");
  const highlight = first<any>(rpr, "highlight");
  const underline = first<any>(rpr, "u");
  const fonts = first<any>(rpr, "rFonts");
  const parsed: Record<string, any> = {};
  if (first<any>(rpr, "b") !== undefined) {
    parsed.bold = true;
  }
  if (first<any>(rpr, "i") !== undefined) {
    parsed.italic = true;
  }
  if (underline !== undefined && String(underline.val || "single") !== "none") {
    parsed.underline = true;
  }
  if (color?.val && String(color.val).toLowerCase() !== "auto") {
    parsed.color = `#${String(color.val).replace(/^#/, "").toUpperCase()}`;
  }
  if (size?.val && !Number.isNaN(Number(size.val))) {
    parsed.font_size_pt = Number(size.val) / 2;
  }
  if (highlight?.val && String(highlight.val).toLowerCase() !== "none") {
    parsed.highlight = String(highlight.val);
  }
  if (fonts?.ascii || fonts?.eastAsia || fonts?.hAnsi) {
    parsed.font = String(fonts.eastAsia || fonts.ascii || fonts.hAnsi);
  }
  return parsed;
}

function sameRunFormat(left: Record<string, any>, right: Record<string, any>): boolean {
  const keys = ["bold", "italic", "underline", "color", "font_size_pt", "highlight", "font"];
  return keys.every((key) => left[key] === right[key]);
}

function parseRuns(paragraph: XmlNode, styles: Record<string, StyleInfo>): Record<string, any>[] {
  const style = paragraphStyle(paragraph);
  const baseRun = styles[style]?.run || {};
  const runs: Record<string, any>[] = [];
  for (const run of asArray<any>(paragraph.r)) {
    const runStyle = first<any>(first<any>(run, "rPr"), "rStyle");
    const runStyleProps = runStyle?.val ? styles[String(runStyle.val)]?.run || {} : {};
    const props = {
      ...baseRun,
      ...runStyleProps,
      ...parseRunProperties(first<any>(run, "rPr"))
    };
    const parts: string[] = [];
    for (const [key, value] of childEntries(run)) {
      if (key === "t") {
        parts.push(textContent(value));
      } else if (key === "tab") {
        parts.push("\t");
      } else if (key === "br" || key === "cr") {
        parts.push("\n");
      }
    }
    const text = parts.join("");
    if (text) {
      const current = { text, ...props };
      const previous = runs[runs.length - 1];
      if (previous && sameRunFormat(previous, current)) {
        previous.text = `${previous.text}${text}`;
      } else {
        runs.push(current);
      }
    }
  }
  return runs;
}

function directParagraphText(paragraph: XmlNode): string {
  const parts: string[] = [];
  for (const run of asArray<any>(paragraph.r)) {
    for (const [key, value] of childEntries(run)) {
      if (key === "t") {
        parts.push(textContent(value));
      } else if (key === "tab") {
        parts.push("\t");
      } else if (key === "br" || key === "cr") {
        parts.push("\n");
      }
    }
  }
  return parts.join("").trim();
}

function allText(element: any): string {
  return findAll(element, "t").map(textContent).join("").trim();
}

function extractObjectLabels(text: string): string[] {
  const labels = text.match(/オブジェクト[^オブジェクト\s]+/g) || [];
  return labels.length >= 2 ? labels : text ? [text] : [];
}

function chartText(node: any): string {
  return findAll(node, "t").map(textContent).join("").trim();
}

function chartPointText(point: any): string {
  const v = first<any>(point, "v");
  return textContent(v || point).trim();
}

function chartCacheValues(node: any): string[] {
  const cache = first<any>(node, "strCache") || first<any>(node, "numCache") || node;
  return asArray<any>(cache?.pt).map(chartPointText).filter((value) => value.length > 0);
}

function chartAxisValues(node: any): string[] {
  const ref = first<any>(node, "strRef") || first<any>(node, "numRef") || node;
  return chartCacheValues(ref);
}

function parseChartSeries(chartTypeNode: XmlNode): Record<string, any>[] {
  return asArray<any>(chartTypeNode.ser).map((ser, index) => {
    const name = chartText(first<any>(ser, "tx")) || `series${index + 1}`;
    const categories = chartAxisValues(first<any>(ser, "cat") || {});
    const values = chartAxisValues(first<any>(ser, "val") || {}).map((value) => Number(value));
    return {
      name,
      categories,
      values: values.filter((value) => !Number.isNaN(value))
    };
  });
}

function mermaidForChart(chart: Record<string, any>): Record<string, string> | undefined {
  const series = asArray<Record<string, any>>(chart.series)[0];
  const categories = asArray<string>(chart.categories);
  const values = asArray<number>(series?.values);
  if (!series || !categories.length || !values.length) {
    return undefined;
  }
  if (chart.chart_type === "pie") {
    const lines = [`pie title ${chart.title || "Chart"}`];
    categories.forEach((category, index) => {
      lines.push(`  "${String(category).replace(/"/g, '\\"')}" : ${values[index] || 0}`);
    });
    return { type: "pie", code: lines.join("\n") };
  }
  if (["bar", "line", "area"].includes(String(chart.chart_type))) {
    const max = Math.max(...values, 0);
    const mark = chart.chart_type === "line" ? "line" : "bar";
    return {
      type: "xychart-beta",
      code: [
        "xychart-beta",
        `  title "${String(chart.title || "Chart").replace(/"/g, '\\"')}"`,
        `  x-axis [${categories.map((category) => `"${String(category).replace(/"/g, '\\"')}"`).join(", ")}]`,
        `  y-axis "${String(series.name || "value").replace(/"/g, '\\"')}" 0 --> ${Math.max(max, 1)}`,
        `  ${mark} [${values.join(", ")}]`
      ].join("\n")
    };
  }
  return undefined;
}

function dataTableForChart(chart: Record<string, any>): string[][] {
  const rows = [["chart", "series", "category", "value"]];
  for (const series of asArray<Record<string, any>>(chart.series)) {
    const categories = asArray<string>(series.categories?.length ? series.categories : chart.categories);
    const values = asArray<number>(series.values);
    values.forEach((value, index) => {
      rows.push([chart.title || "Chart", series.name || "", categories[index] || String(index + 1), String(value)]);
    });
  }
  return rows;
}

function parseChart(zip: AdmZip, target: string): Record<string, any> | undefined {
  if (!hasEntry(zip, target)) {
    return undefined;
  }
  const root = readXml(zip, target);
  const chart = root.chartSpace?.chart;
  const plotArea = chart?.plotArea;
  if (!plotArea) {
    return undefined;
  }
  const chartTypeMap: Record<string, string> = {
    barChart: "bar",
    lineChart: "line",
    pieChart: "pie",
    areaChart: "area",
    scatterChart: "scatter"
  };
  const chartTypeKey = Object.keys(chartTypeMap).find((key) => plotArea[key]);
  if (!chartTypeKey) {
    return undefined;
  }
  const chartTypeNode = first<any>(plotArea, chartTypeKey) || plotArea[chartTypeKey];
  const series = parseChartSeries(chartTypeNode);
  const parsed: Record<string, any> = {
    type: "chart",
    chart_type: chartTypeMap[chartTypeKey],
    title: chartText(chart?.title) || "Chart",
    categories: series[0]?.categories || [],
    series,
    source: target
  };
  parsed.mermaid_candidate = mermaidForChart(parsed);
  parsed.data_table = dataTableForChart(parsed);
  return parsed;
}

function parseDrawing(drawing: XmlNode, ctx: Context, zip: AdmZip): Drawing[] {
  const docPr = first<any>(drawing, "docPr") || findAll(drawing, "docPr")[0];
  const extent = first<any>(drawing, "extent") || findAll(drawing, "extent")[0];
  const base: Drawing = {
    doc_pr_id: docPr?.id,
    name: docPr?.name,
    description: docPr?.descr
  };
  if (extent) {
    const cx = Number(extent.cx || 0);
    const cy = Number(extent.cy || 0);
    base.size = {
      cx_emu: cx,
      cy_emu: cy,
      width_px: Math.round((cx / EMU_PER_PIXEL) * 100) / 100,
      height_px: Math.round((cy / EMU_PER_PIXEL) * 100) / 100
    };
  }

  const elements: Drawing[] = [];
  for (const chartRef of findAll(drawing, "chart")) {
    const rid = chartRef.id;
    const rel = rid ? ctx.rels[String(rid)] : undefined;
    if (!rid || !rel) {
      continue;
    }
    const chart = parseChart(zip, rel.target);
    if (chart) {
      elements.push({
        ...base,
        ...chart,
        relationship_id: rid
      });
    }
  }

  for (const blip of findAll(drawing, "blip")) {
    const rid = blip.embed;
    const rel = rid ? ctx.rels[String(rid)] : undefined;
    if (!rid || !rel) {
      continue;
    }
    elements.push({
      ...base,
      type: "image",
      relationship_id: rid,
      resource: copyResource(ctx, zip, rel.target)
    });
  }

  const shapeText = allText(drawing);
  if (shapeText && elements.length === 0) {
    const labels = extractObjectLabels(shapeText);
    const mermaid = labels.length >= 2
      ? {
          type: "flowchart_lr",
          code: ["flowchart LR", `  A[${labels[0]}] --> B[${labels[1]}]`].join("\n"),
          note: "図形内テキストから推定した候補。矢印や配置は docx XML だけでは確定できない。"
        }
      : undefined;
    elements.push({
      ...base,
      type: "shape_text",
      text: shapeText,
      labels,
      mermaid_candidate: mermaid
    });
  }

  return elements;
}

function parseTable(table: XmlNode): Block {
  const rows: string[][] = [];
  for (const tr of asArray<any>(table.tr)) {
    const row: string[] = [];
    for (const tc of asArray<any>(tr.tc)) {
      row.push(allText(tc));
    }
    rows.push(row);
  }
  return { type: "table", rows };
}

function parseDocx(sourceDocx: string, outputDir: string): Record<string, any> {
  const resourceDir = path.join(outputDir, `${path.parse(sourceDocx).name}resource`);
  fs.rmSync(resourceDir, { recursive: true, force: true });
  fs.mkdirSync(resourceDir, { recursive: true });

  const zip = new AdmZip(sourceDocx);
  const ctx: Context = {
    sourceDocx,
    outputDir,
    resourceDir,
    rels: parseRelationships(zip),
    styles: parseStyles(zip),
    numbering: parseNumbering(zip),
    copiedResources: {}
  };
  const document = readXml(zip, "word/document.xml");
  const body = document.document?.body;
  if (!body) {
    throw new Error("word/document.xml に body がありません");
  }

  const blocks: Block[] = [];
  const images: Drawing[] = [];
  const shapes: Drawing[] = [];
  const charts: Drawing[] = [];
  let index = 0;

  const appendBlocks = (container: XmlNode) => {
    for (const [key, value] of childEntries(container)) {
      for (const child of asArray<any>(value)) {
        if (key === "p") {
          const drawings: Drawing[] = [];
          for (const drawing of findAll(child, "drawing")) {
            drawings.push(...parseDrawing(drawing, ctx, zip));
          }
          for (const drawing of drawings) {
            if (drawing.type === "image") {
              images.push(drawing);
            } else if (drawing.type === "chart") {
              charts.push(drawing);
            } else {
              shapes.push(drawing);
            }
          }
          blocks.push({
            type: "paragraph",
            index,
            style: paragraphStyle(child),
            text: directParagraphText(child),
            runs: parseRuns(child, ctx.styles),
            paragraph_format: paragraphFormat(child, ctx.styles),
            numbering: paragraphNumbering(child, ctx.numbering),
            drawings
          });
        } else if (key === "tbl") {
          const table = parseTable(child);
          table.index = index;
          blocks.push(table);
        } else if (child && typeof child === "object") {
          appendBlocks(child);
        }
        index += 1;
      }
    }
  };
  appendBlocks(body);

  return {
    document: {
      source: sourceDocx,
      format: "docx"
    },
    blocks,
    images,
    shapes,
    charts,
    conversion_notes: [
      "段落/見出し/画像/表/チャートは Markdown へ直接変換する。",
      "文字装飾はAIが読み取れる本文を保ちながら、HTML混在Markdownで補足する。",
      "Word 図形は DrawingML の座標や形状情報が Markdown と対応しづらいため、Mermaid 化または画像化を検討対象にする。",
      "本 POC は docx の zip/OpenXML を直接解析している。"
    ]
  };
}

function mdEscapeTableCell(value: string): string {
  return htmlEscape(value).replace(/\|/g, "\\|").replace(/\n/g, "<br>");
}

function htmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function styleAttr(styles: Record<string, string | number | undefined>): string {
  return Object.entries(styles)
    .filter(([, value]) => value !== undefined && value !== "")
    .map(([key, value]) => `${key}:${value}`)
    .join("; ");
}

function markdownTable(rows: string[][]): string {
  if (!rows.length) {
    return "";
  }
  const width = Math.max(...rows.map((row) => row.length));
  const normalized = rows.map((row) => [...row, ...Array(width - row.length).fill("")]);
  const header = `| ${normalized[0].map(mdEscapeTableCell).join(" | ")} |`;
  const sep = `| ${Array(width).fill("---").join(" | ")} |`;
  const body = normalized.slice(1).map((row) => `| ${row.map(mdEscapeTableCell).join(" | ")} |`);
  return [header, sep, ...body].join("\n");
}

function cssHighlightColor(value: string): string {
  const colors: Record<string, string> = {
    black: "black",
    blue: "blue",
    cyan: "cyan",
    green: "lime",
    magenta: "magenta",
    red: "red",
    yellow: "yellow",
    white: "white",
    darkBlue: "navy",
    darkCyan: "teal",
    darkGreen: "green",
    darkMagenta: "purple",
    darkRed: "maroon",
    darkYellow: "olive",
    darkGray: "gray",
    lightGray: "lightgray"
  };
  return colors[value] || value;
}

function renderRun(run: Record<string, any>): string {
  const text = String(run.text || "");
  const styles = styleAttr({
    color: run.color,
    "font-size": run.font_size_pt ? `${run.font_size_pt}pt` : undefined,
    "background-color": run.highlight ? cssHighlightColor(String(run.highlight)) : undefined
  });
  const needsHtml = Boolean(styles || run.underline);
  let rendered = htmlEscape(text);

  if (styles) {
    rendered = `<span style="${styles}">${rendered}</span>`;
  }
  if (run.underline) {
    rendered = `<u>${rendered}</u>`;
  }
  if (run.bold) {
    rendered = needsHtml ? `<strong>${rendered}</strong>` : `**${rendered}**`;
  }
  if (run.italic) {
    rendered = needsHtml ? `<em>${rendered}</em>` : `*${rendered}*`;
  }
  return rendered;
}

function renderRuns(block: Block): string {
  const runs = asArray<Record<string, any>>(block.runs);
  if (!runs.length) {
    return htmlEscape(String(block.text || ""));
  }
  return runs.map(renderRun).join("").trim();
}

function renderParagraphText(block: Block, includeParagraphStyle = true): string {
  const rendered = renderRuns(block);
  const align = String(block.paragraph_format?.align || "");
  const cssAlign: Record<string, string> = {
    center: "center",
    right: "right",
    both: "justify",
    distribute: "justify"
  };
  if (!includeParagraphStyle || !rendered || !cssAlign[align]) {
    return rendered;
  }
  return `<p style="text-align:${cssAlign[align]}">${rendered}</p>`;
}

function renderListPrefix(block: Block): string {
  const numbering = block.numbering;
  if (!numbering) {
    return "";
  }
  const indent = "  ".repeat(Number(numbering.ilvl || 0));
  if (String(numbering.format) === "bullet") {
    const marker = String(numbering.text || "-");
    const normalizedMarker = ["", "", "●", "•", "o"].includes(marker) ? "-" : marker;
    return `${indent}${normalizedMarker} `;
  }
  const marker = String(numbering.text || "1.").replace(/%\d+/g, "1");
  return `${indent}${marker} `;
}

function visibleImageDrawings(block: Block): Drawing[] {
  return (block.drawings || []).filter((drawing: Drawing) => {
    if (drawing.type !== "image") {
      return false;
    }
    const resource = drawing.resource || {};
    return !(Number(resource.width_px || 0) <= 1 && Number(resource.height_px || 0) <= 1);
  });
}

function markdownForBlock(block: Block): string[] {
  if (block.type === "table") {
    const tableMd = markdownTable(block.rows || []);
    return tableMd ? [tableMd] : [];
  }

  const text = String(block.text || "");
  const style = String(block.style || "Normal");
  const lines: string[] = [];
  if (text) {
    if (style === "Title") {
      const renderedText = renderParagraphText(block, false);
      lines.push(`# ${renderedText}`);
    } else if (style === "Subtitle") {
      const renderedText = renderParagraphText(block, false);
      lines.push(`*${renderedText}*`);
    } else if (style.startsWith("Heading")) {
      const renderedText = renderParagraphText(block, false);
      const level = Math.min(Math.max(Number(style.replace(/[^0-9]/g, "") || "1") + 1, 2), 6);
      lines.push(`${"#".repeat(level)} ${renderedText}`);
    } else {
      const renderedText = renderParagraphText(block);
      lines.push(`${renderListPrefix(block)}${renderedText}`);
    }
  }

  const visibleImages = visibleImageDrawings(block);
  for (const drawing of block.drawings || []) {
    if (drawing.type === "image") {
      if (!visibleImages.includes(drawing)) {
        continue;
      }
      const resource = drawing.resource;
      const alt = htmlEscape(String(drawing.description || drawing.name || path.basename(resource.path))).replace(/"/g, "&quot;");
      const width = Number(drawing.size?.width_px || 0);
      const height = Number(drawing.size?.height_px || 0);
      if (width > 0 && height > 0) {
        lines.push(`<img src="${resource.path}" alt="${alt}" width="${Math.round(width)}" height="${Math.round(height)}">`);
      } else {
        lines.push(`![${alt}](${resource.path})`);
      }
    } else if (drawing.type === "shape_text") {
      const mermaid = drawing.mermaid_candidate;
      if (mermaid) {
        lines.push("```mermaid", mermaid.code, "```");
      } else {
        lines.push(`> 図形テキスト: ${drawing.text || ""}`);
      }
    } else if (drawing.type === "chart") {
      const mermaid = drawing.mermaid_candidate;
      if (mermaid?.code) {
        lines.push("```mermaid", mermaid.code, "```");
      } else if (drawing.data_table) {
        const tableMd = markdownTable(drawing.data_table);
        if (tableMd) {
          lines.push(tableMd);
        }
      } else {
        lines.push(`> チャート: ${drawing.title || drawing.name || "未対応チャート"}`);
      }
    }
  }
  return lines;
}

function buildMarkdown(intermediate: Record<string, any>): string {
  const chunks: string[] = [];
  for (const block of intermediate.blocks || []) {
    const lines = markdownForBlock(block);
    if (lines.length) {
      chunks.push(lines.join("\n"));
    }
  }
  return `${chunks.join("\n\n").trimEnd()}\n`;
}
