import * as fs from "fs";
import * as path from "path";
import AdmZip = require("adm-zip");
import {
  asArray,
  childEntries,
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

type Dict = Record<string, any>;

type Relationship = {
  type: string;
  target: string;
};

type SlideSize = {
  width: number;
  height: number;
};

type Transform = {
  x: number;
  y: number;
  width: number;
  height: number;
  chX: number;
  chY: number;
  chWidth: number;
  chHeight: number;
};

type SlideItem = {
  id: string;
  name: string;
  kind: "shape" | "text" | "image" | "line" | "table";
  x: number;
  y: number;
  width: number;
  height: number;
  text?: string;
  fill?: string;
  stroke?: string;
  fontSize?: number;
  fontColor?: string;
  align?: string;
  html?: string;
  imageDataUri?: string;
};

type Slide = {
  name: string;
  path: string;
  items: SlideItem[];
};

export async function convertPptxToMarp(sourcePath: string): Promise<string> {
  const sourceName = path.parse(sourcePath).name;
  const outputDir = path.join(path.dirname(sourcePath), sourceName);
  const slideDir = path.join(outputDir, "slide");
  fs.rmSync(slideDir, { recursive: true, force: true });
  fs.mkdirSync(slideDir, { recursive: true });

  const zip = new AdmZip(sourcePath);
  const presentation = readXml(zip, "ppt/presentation.xml").presentation;
  const slideSize = parseSlideSize(presentation?.sldSz);
  const slides = parseSlides(zip);

  const mdLines = ["---", "marp: true", "---", ""];
  slides.forEach((slide, index) => {
    const slideNumber = index + 1;
    const svgName = `slide${String(slideNumber).padStart(3, "0")}.drawio.svg`;
    fs.writeFileSync(
      path.join(slideDir, svgName),
      makeDrawioSvg(slide, slideNumber, slideSize),
      "utf8"
    );
    if (index > 0) {
      mdLines.push("---", "");
    }
    mdLines.push(`![bg contain](slide/${svgName})`, "");
  });

  const markdownPath = path.join(outputDir, `${sourceName}.md`);
  fs.writeFileSync(markdownPath, mdLines.join("\n"), "utf8");
  return markdownPath;
}

function emuToPx(value: unknown): number {
  return round((Number(value || 0) / EMU_PER_INCH) * PX_PER_INCH);
}

function parseSlideSize(node: Dict | undefined): SlideSize {
  return {
    width: emuToPx(node?.cx || 12192000),
    height: emuToPx(node?.cy || 6858000)
  };
}

function readRelationships(zip: AdmZip, name: string): Record<string, Relationship> {
  if (!hasEntry(zip, name)) {
    return {};
  }
  const root = readXml(zip, name);
  const rels: Record<string, Relationship> = {};
  for (const rel of asArray<any>(root.Relationships?.Relationship)) {
    rels[String(rel.Id)] = {
      type: String(rel.Type || ""),
      target: String(rel.Target || "")
    };
  }
  return rels;
}

function parseSlides(zip: AdmZip): Slide[] {
  const presentation = readXml(zip, "ppt/presentation.xml").presentation;
  const rels = readRelationships(zip, "ppt/_rels/presentation.xml.rels");
  const slides: Slide[] = [];
  for (const slideId of asArray<any>(presentation?.sldIdLst?.sldId)) {
    const relId = String(slideId.id || "");
    const target = rels[relId]?.target;
    if (!target) {
      continue;
    }
    const slidePath = resolvePackagePath("ppt/presentation.xml", target);
    const slideRels = readRelationships(zip, `${path.posix.dirname(slidePath)}/_rels/${path.posix.basename(slidePath)}.rels`);
    const root = readXml(zip, slidePath).sld;
    slides.push({
      name: path.basename(slidePath, ".xml"),
      path: slidePath,
      items: parseSlideItems(zip, root?.cSld?.spTree, slidePath, slideRels)
    });
  }
  return slides;
}

function parseSlideItems(zip: AdmZip, spTree: Dict | undefined, slidePath: string, rels: Record<string, Relationship>): SlideItem[] {
  const items: SlideItem[] = [];
  const identity: Transform = {
    x: 0,
    y: 0,
    width: 1,
    height: 1,
    chX: 0,
    chY: 0,
    chWidth: 1,
    chHeight: 1
  };
  visitShapeContainer(zip, spTree, slidePath, rels, identity, items);
  return items;
}

function visitShapeContainer(
  zip: AdmZip,
  container: Dict | undefined,
  slidePath: string,
  rels: Record<string, Relationship>,
  parent: Transform,
  items: SlideItem[]
): void {
  if (!container) {
    return;
  }
  for (const shape of asArray<any>(container.sp)) {
    const item = parseShape(zip, shape, slidePath, rels, parent);
    if (item) {
      items.push(item);
    }
  }
  for (const picture of asArray<any>(container.pic)) {
    const item = parseShape(zip, picture, slidePath, rels, parent);
    if (item) {
      items.push(item);
    }
  }
  for (const frame of asArray<any>(container.graphicFrame)) {
    const item = parseGraphicFrame(frame, parent);
    if (item) {
      items.push(item);
    }
  }
  for (const group of asArray<any>(container.grpSp)) {
    const groupTransform = absoluteTransform(parseTransform(group.grpSpPr?.xfrm), parent);
    visitShapeContainer(zip, group, slidePath, rels, groupTransform, items);
  }
}

function parseGraphicFrame(frame: Dict, parent: Transform): SlideItem | undefined {
  const table = findAll(frame, "tbl")[0];
  if (!table) {
    return undefined;
  }
  const nv = findAll(frame, "cNvPr")[0] || {};
  const transform = absoluteTransform(parseTransform(findAll(frame, "xfrm")[0]), parent);
  return {
    id: String(nv.id || `table-${itemsafeName(String(nv.name || ""))}`),
    name: String(nv.name || ""),
    kind: "table",
    x: transform.x,
    y: transform.y,
    width: transform.width,
    height: transform.height,
    html: tableToHtml(table)
  };
}

function parseShape(zip: AdmZip, shape: Dict, slidePath: string, rels: Record<string, Relationship>, parent: Transform): SlideItem | undefined {
  const nv = findAll(shape, "cNvPr")[0] || {};
  const xfrm = findAll(shape, "xfrm")[0];
  const transform = absoluteTransform(parseTransform(xfrm), parent);
  const text = shapeText(shape);
  const blip = findAll(shape, "blip")[0];
  const relId = String(blip?.embed || "");
  const image = relId ? imageDataUri(zip, slidePath, rels[relId]?.target || "") : "";
  const preset = String(findAll(shape, "prstGeom")[0]?.prst || "");
  const stroke = findColor(first<any>(findAll(shape, "ln")[0], "solidFill")) || findColor(findAll(shape, "ln")[0]);
  const fill = findColor(findAll(shape, "solidFill")[0]);
  const common = {
    id: String(nv.id || `item-${itemsafeName(String(nv.name || ""))}`),
    name: String(nv.name || ""),
    x: transform.x,
    y: transform.y,
    width: transform.width,
    height: transform.height,
    fill,
    stroke
  };

  if (image) {
    return { ...common, kind: "image", imageDataUri: image };
  }
  if (preset === "line" || transform.height === 0 || transform.width === 0) {
    return { ...common, kind: "line" };
  }
  if (text) {
    const run = firstRunProperties(shape);
    return {
      ...common,
      kind: "text",
      text,
      fontSize: run.fontSize,
      fontColor: run.color,
      align: paragraphAlign(shape)
    };
  }
  if (fill || stroke) {
    return { ...common, kind: "shape" };
  }
  return undefined;
}

function parseTransform(xfrm: Dict | undefined): Transform {
  return {
    x: emuToPx(xfrm?.off?.x),
    y: emuToPx(xfrm?.off?.y),
    width: emuToPx(xfrm?.ext?.cx),
    height: emuToPx(xfrm?.ext?.cy),
    chX: emuToPx(xfrm?.chOff?.x),
    chY: emuToPx(xfrm?.chOff?.y),
    chWidth: emuToPx(xfrm?.chExt?.cx) || 1,
    chHeight: emuToPx(xfrm?.chExt?.cy) || 1
  };
}

function absoluteTransform(child: Transform, parent: Transform): Transform {
  const scaleX = parent.width / parent.chWidth;
  const scaleY = parent.height / parent.chHeight;
  return {
    x: round(parent.x + (child.x - parent.chX) * scaleX),
    y: round(parent.y + (child.y - parent.chY) * scaleY),
    width: round(child.width * scaleX),
    height: round(child.height * scaleY),
    chX: child.chX,
    chY: child.chY,
    chWidth: child.chWidth,
    chHeight: child.chHeight
  };
}

function shapeText(shape: Dict): string {
  return asArray<any>(shape.txBody?.p)
    .map((paragraph) => findAll(paragraph, "t").map(textContent).join(""))
    .filter((text) => text.trim())
    .join("\n");
}

function firstRunProperties(shape: Dict): { fontSize?: number; color?: string } {
  const rPr = findAll(shape, "rPr")[0];
  const size = Number(rPr?.sz || 0);
  return {
    fontSize: size ? round(size / 100) : undefined,
    color: findColor(rPr?.solidFill)
  };
}

function paragraphAlign(shape: Dict): string | undefined {
  const align = String(findAll(shape, "pPr")[0]?.algn || "");
  return align || undefined;
}

function findColor(node: Dict | undefined): string | undefined {
  const color = first<any>(node, "srgbClr") || node?.srgbClr;
  return color?.val ? `#${String(color.val)}` : undefined;
}

function imageDataUri(zip: AdmZip, basePath: string, target: string): string {
  if (!target) {
    return "";
  }
  const imagePath = resolvePackagePath(basePath, target);
  const data = readBuffer(zip, imagePath);
  const suffix = path.extname(imagePath).toLowerCase().replace(/^\./, "");
  const mime = suffix === "jpg" || suffix === "jpeg" ? "image/jpeg" : suffix === "svg" ? "image/svg+xml" : `image/${suffix || "octet-stream"}`;
  return `data:${mime};base64,${data.toString("base64")}`;
}

function itemsafeName(value: string): string {
  return value.replace(/[^A-Za-z0-9_-]/g, "-") || "shape";
}

function makeDrawioSvg(slide: Slide, slideNumber: number, size: SlideSize): string {
  const drawioXml = makeDrawioXml(slide, slideNumber, size);
  const lines = [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<svg xmlns="http://www.w3.org/2000/svg" width=${quoteAttr(String(size.width))} height=${quoteAttr(String(size.height))} viewBox=${quoteAttr(`0 0 ${size.width} ${size.height}`)} content=${quoteAttr(drawioXml)}>`
  ];
  lines.push(`<rect x="0" y="0" width=${quoteAttr(String(size.width))} height=${quoteAttr(String(size.height))} fill="#ffffff"/>`);
  for (const item of slide.items) {
    lines.push(svgElement(item));
  }
  lines.push("</svg>");
  return lines.join("\n");
}

function svgElement(item: SlideItem): string {
  if (item.kind === "image" && item.imageDataUri) {
    return `<image x=${quoteAttr(String(item.x))} y=${quoteAttr(String(item.y))} width=${quoteAttr(String(item.width || 1))} height=${quoteAttr(String(item.height || 1))} href=${quoteAttr(item.imageDataUri)} preserveAspectRatio="none"/>`;
  }
  if (item.kind === "line") {
    const x2 = item.x + item.width;
    const y2 = item.y + item.height;
    return `<line x1=${quoteAttr(String(item.x))} y1=${quoteAttr(String(item.y))} x2=${quoteAttr(String(x2))} y2=${quoteAttr(String(y2))} stroke=${quoteAttr(item.stroke || "#000000")} stroke-width="2"/>`;
  }
  if (item.kind === "text") {
    return `<foreignObject x=${quoteAttr(String(item.x))} y=${quoteAttr(String(item.y))} width=${quoteAttr(String(item.width || 1))} height=${quoteAttr(String(item.height || 1))}><div xmlns="http://www.w3.org/1999/xhtml" style=${quoteAttr(textStyle(item))}>${textToHtml(item.text || "")}</div></foreignObject>`;
  }
  if (item.kind === "table") {
    return `<foreignObject x=${quoteAttr(String(item.x))} y=${quoteAttr(String(item.y))} width=${quoteAttr(String(item.width || 1))} height=${quoteAttr(String(item.height || 1))}><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%">${item.html || ""}</div></foreignObject>`;
  }
  return `<rect x=${quoteAttr(String(item.x))} y=${quoteAttr(String(item.y))} width=${quoteAttr(String(item.width || 1))} height=${quoteAttr(String(item.height || 1))} fill=${quoteAttr(item.fill || "none")} stroke=${quoteAttr(item.stroke || "none")}/>`;
}

function tableToHtml(table: Dict): string {
  const rows = asArray<any>(table.tr).map((row) =>
    asArray<any>(row.tc).map((cell) => findAll(cell, "t").map(textContent).join(""))
  );
  const htmlRows = rows.map((row) => {
    const cells = row.map((cell) => `<td style="border:1px solid #666;padding:4px;text-align:center;vertical-align:middle;">${xmlEscape(cell)}</td>`);
    return `<tr>${cells.join("")}</tr>`;
  });
  return `<table style="border-collapse:collapse;width:100%;height:100%;font-family:Arial,sans-serif;font-size:14px;table-layout:fixed;">${htmlRows.join("")}</table>`;
}

function textStyle(item: SlideItem): string {
  const align = item.align === "ctr" ? "center" : item.align === "r" ? "right" : "left";
  return [
    "width:100%",
    "height:100%",
    "box-sizing:border-box",
    "white-space:pre-wrap",
    "overflow:hidden",
    "font-family:Arial,sans-serif",
    `font-size:${item.fontSize || 18}px`,
    `color:${item.fontColor || "#000000"}`,
    `text-align:${align}`
  ].join(";");
}

function textToHtml(text: string): string {
  return xmlEscape(text).replace(/\n/g, "<br/>");
}

function makeDrawioXml(slide: Slide, slideNumber: number, size: SlideSize): string {
  const lines = [
    `<mxfile host="office-to-mdrow">`,
    `  <diagram id=${quoteAttr(`slide-${slideNumber}`)} name=${quoteAttr(slide.name)}>`,
    `    <mxGraphModel dx=${quoteAttr(String(size.width))} dy=${quoteAttr(String(size.height))} grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth=${quoteAttr(String(size.width))} pageHeight=${quoteAttr(String(size.height))}>`,
    "      <root>",
    '        <mxCell id="0"/>',
    '        <mxCell id="1" parent="0"/>'
  ];
  for (const item of slide.items) {
    lines.push(drawioCell(item));
  }
  lines.push("      </root>", "    </mxGraphModel>", "  </diagram>", "</mxfile>");
  return lines.join("\n");
}

function drawioCell(item: SlideItem): string {
  const id = `item-${item.id}`;
  const value = item.kind === "text" ? String(item.text || "") : item.kind === "table" ? String(item.html || "") : "";
  const style = drawioStyle(item);
  return [
    `        <mxCell id=${quoteAttr(id)} value=${quoteAttr(value)} style=${quoteAttr(style)} vertex="1" parent="1">`,
    `          <mxGeometry x=${quoteAttr(String(item.x))} y=${quoteAttr(String(item.y))} width=${quoteAttr(String(item.width || 1))} height=${quoteAttr(String(item.height || 1))} as="geometry"/>`,
    "        </mxCell>"
  ].join("\n");
}

function drawioStyle(item: SlideItem): string {
  if (item.kind === "image") {
    return `shape=image;html=1;imageAspect=0;aspect=fixed;image=${item.imageDataUri || ""};`;
  }
  if (item.kind === "text") {
    return `text;html=1;whiteSpace=wrap;strokeColor=none;fillColor=none;fontSize=${item.fontSize || 18};fontColor=${item.fontColor || "#000000"};`;
  }
  if (item.kind === "table") {
    return "html=1;whiteSpace=wrap;overflow=fill;rounded=0;fillColor=none;strokeColor=none;";
  }
  if (item.kind === "line") {
    return `shape=line;html=1;strokeColor=${item.stroke || "#000000"};`;
  }
  return `rounded=0;whiteSpace=wrap;html=1;fillColor=${item.fill || "none"};strokeColor=${item.stroke || "none"};`;
}
