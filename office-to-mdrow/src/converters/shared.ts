import * as path from "path";
import AdmZip = require("adm-zip");
import { XMLParser } from "fast-xml-parser";

export type XmlNode = Record<string, any>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true,
  textNodeName: "#text"
});

export function parseXml(text: string): XmlNode {
  return parser.parse(text);
}

export function readXml(zip: AdmZip, name: string): XmlNode {
  const entry = zip.getEntry(name);
  if (!entry) {
    throw new Error(`Office package entry not found: ${name}`);
  }
  return parseXml(entry.getData().toString("utf8"));
}

export function readText(zip: AdmZip, name: string): string {
  const entry = zip.getEntry(name);
  if (!entry) {
    throw new Error(`Office package entry not found: ${name}`);
  }
  return entry.getData().toString("utf8");
}

export function readBuffer(zip: AdmZip, name: string): Buffer {
  const entry = zip.getEntry(name);
  if (!entry) {
    throw new Error(`Office package entry not found: ${name}`);
  }
  return entry.getData();
}

export function hasEntry(zip: AdmZip, name: string): boolean {
  return Boolean(zip.getEntry(name));
}

export function entries(zip: AdmZip): string[] {
  return zip.getEntries().map((entry: AdmZip.IZipEntry) => entry.entryName);
}

export function asArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

export function first<T = any>(node: XmlNode | undefined, key: string): T | undefined {
  if (!node) {
    return undefined;
  }
  return asArray<T>(node[key])[0];
}

export function localName(tag: string): string {
  return tag.includes(":") ? tag.split(":").pop() || tag : tag;
}

export function childEntries(node: any): Array<[string, any]> {
  if (!node || typeof node !== "object") {
    return [];
  }
  return Object.entries(node).filter(([key]) => !key.startsWith("#") && key !== ":@");
}

export function findAll(node: any, key: string): any[] {
  const found: any[] = [];
  const visit = (current: any) => {
    if (!current || typeof current !== "object") {
      return;
    }
    for (const [childKey, value] of childEntries(current)) {
      if (childKey === key) {
        found.push(...asArray(value));
      }
      for (const item of asArray(value)) {
        visit(item);
      }
    }
  };
  visit(node);
  return found;
}

export function textContent(node: any): string {
  if (node === undefined || node === null) {
    return "";
  }
  if (typeof node === "string" || typeof node === "number" || typeof node === "boolean") {
    return String(node);
  }
  if (Array.isArray(node)) {
    return node.map(textContent).join("");
  }
  let text = typeof node["#text"] === "string" ? node["#text"] : "";
  for (const [, value] of childEntries(node)) {
    text += textContent(value);
  }
  return text;
}

export function resolvePackagePath(baseFile: string, target: string): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  return path.posix.normalize(path.posix.join(path.posix.dirname(baseFile), target));
}

export function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function quoteAttr(value: string): string {
  return `"${xmlEscape(value).replace(/"/g, "&quot;")}"`;
}

export function round(value: number, digits = 2): number {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}
