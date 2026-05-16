import * as fs from "fs";
import * as path from "path";
import AdmZip = require("adm-zip");
import { asArray, childEntries, findAll, first, readBuffer, readXml, textContent, XmlNode } from "./shared";

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
  copiedResources: Record<string, string>;
};

type Drawing = Record<string, any>;
type Block = Record<string, any>;

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
  let index = 0;
  for (const [key, value] of childEntries(body)) {
    for (const child of asArray<any>(value)) {
      if (key === "p") {
        const drawings: Drawing[] = [];
        for (const drawing of findAll(child, "drawing")) {
          drawings.push(...parseDrawing(drawing, ctx, zip));
        }
        for (const drawing of drawings) {
          if (drawing.type === "image") {
            images.push(drawing);
          } else {
            shapes.push(drawing);
          }
        }
        blocks.push({
          type: "paragraph",
          index,
          style: paragraphStyle(child),
          text: directParagraphText(child),
          drawings
        });
      } else if (key === "tbl") {
        const table = parseTable(child);
        table.index = index;
        blocks.push(table);
      }
      index += 1;
    }
  }

  return {
    document: {
      source: sourceDocx,
      format: "docx"
    },
    blocks,
    images,
    shapes,
    conversion_notes: [
      "段落/見出し/画像/表は Markdown へ直接変換する。",
      "Word 図形は DrawingML の座標や形状情報が Markdown と対応しづらいため、画像化または Mermaid 化を検討対象にする。",
      "本 POC は docx の zip/OpenXML を直接解析している。"
    ]
  };
}

function mdEscapeTableCell(value: string): string {
  return value.replace(/\|/g, "\\|").replace(/\n/g, "<br>");
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
      lines.push(`# ${text}`);
    } else if (style === "Subtitle") {
      lines.push(`*${text}*`);
    } else if (style.startsWith("Heading")) {
      const level = Math.min(Math.max(Number(style.replace(/[^0-9]/g, "") || "1") + 1, 2), 6);
      lines.push(`${"#".repeat(level)} ${text}`);
    } else {
      lines.push(text);
    }
  }

  const visibleImages = visibleImageDrawings(block);
  for (const drawing of block.drawings || []) {
    if (drawing.type === "image") {
      if (!visibleImages.includes(drawing)) {
        continue;
      }
      const resource = drawing.resource;
      const alt = drawing.description || drawing.name || path.basename(resource.path);
      lines.push(`![${alt}](${resource.path})`);
    } else if (drawing.type === "shape_text") {
      const mermaid = drawing.mermaid_candidate;
      if (mermaid) {
        lines.push("```mermaid", mermaid.code, "```");
      } else {
        lines.push(`> 図形テキスト: ${drawing.text || ""}`);
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
