import * as fs from "fs";
import * as path from "path";
import AdmZip = require("adm-zip");
import { XMLParser } from "fast-xml-parser";

export type XmlNode = Record<string, any>;

const MAX_PACKAGE_BYTES = 100 * 1024 * 1024;
const MAX_ZIP_ENTRIES = 5000;
const MAX_ENTRY_BYTES = 50 * 1024 * 1024;
const MAX_XML_ENTRY_BYTES = 20 * 1024 * 1024;
const MAX_IMAGE_BYTES = 25 * 1024 * 1024;

const IMAGE_MIME_BY_EXT: Record<string, string> = {
  ".gif": "image/gif",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".png": "image/png",
  ".webp": "image/webp"
};

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true,
  textNodeName: "#text",
  trimValues: false
});

export function parseXml(text: string): XmlNode {
  return parser.parse(text);
}

export function openOfficeZip(filePath: string): AdmZip {
  const stat = fs.statSync(filePath);
  if (stat.size > MAX_PACKAGE_BYTES) {
    throw new Error(`Office file is too large. Max ${Math.round(MAX_PACKAGE_BYTES / 1024 / 1024)} MB is supported.`);
  }

  const zip = new AdmZip(filePath);
  const zipEntries = zip.getEntries();
  if (zipEntries.length > MAX_ZIP_ENTRIES) {
    throw new Error(`Office package has too many entries. Max ${MAX_ZIP_ENTRIES} entries are supported.`);
  }

  for (const entry of zipEntries) {
    normalizePackageEntryPath(entry.entryName);
    if (!entry.isDirectory && Number(entry.header.size || 0) > MAX_ENTRY_BYTES) {
      throw new Error(`Office package entry is too large: ${entry.entryName}`);
    }
  }
  return zip;
}

function hiddenTempName(finalPath: string, suffix: string): string {
  const safeName = path.basename(finalPath).replace(/[^A-Za-z0-9._-]/g, "_") || "output";
  return path.join(
    path.dirname(finalPath),
    `.${safeName}.${process.pid}.${Date.now()}.${Math.random().toString(36).slice(2)}.${suffix}`
  );
}

function assertExistingOutputKind(finalPath: string, expectedKind: "file" | "directory"): void {
  if (!fs.existsSync(finalPath)) {
    return;
  }
  const stat = fs.lstatSync(finalPath);
  if (stat.isSymbolicLink()) {
    throw new Error(`Refusing to overwrite symbolic link: ${path.basename(finalPath)}`);
  }
  if (expectedKind === "file" && !stat.isFile()) {
    throw new Error(`Output path exists and is not a file: ${path.basename(finalPath)}`);
  }
  if (expectedKind === "directory" && !stat.isDirectory()) {
    throw new Error(`Output path exists and is not a directory: ${path.basename(finalPath)}`);
  }
}

export function createTempOutputDirectory(finalDir: string): string {
  assertExistingOutputKind(finalDir, "directory");
  return fs.mkdtempSync(hiddenTempName(finalDir, "tmpdir-"));
}

export function replaceOutputPath(tempPath: string, finalPath: string, expectedKind: "file" | "directory"): void {
  assertExistingOutputKind(finalPath, expectedKind);
  const backupPath = fs.existsSync(finalPath) ? hiddenTempName(finalPath, "backup") : undefined;
  if (backupPath) {
    fs.renameSync(finalPath, backupPath);
  }

  try {
    fs.renameSync(tempPath, finalPath);
  } catch (error) {
    if (backupPath && fs.existsSync(backupPath) && !fs.existsSync(finalPath)) {
      fs.renameSync(backupPath, finalPath);
    }
    throw error;
  }

  if (backupPath && fs.existsSync(backupPath)) {
    fs.rmSync(backupPath, { recursive: true, force: true });
  }
}

export function discardTempOutput(tempPath: string): void {
  if (fs.existsSync(tempPath)) {
    fs.rmSync(tempPath, { recursive: true, force: true });
  }
}

export function writeFileAtomically(finalPath: string, data: string | Buffer, encoding?: BufferEncoding): void {
  assertExistingOutputKind(finalPath, "file");
  const tempPath = hiddenTempName(finalPath, "tmpfile");
  if (typeof data === "string") {
    fs.writeFileSync(tempPath, data, encoding || "utf8");
  } else {
    fs.writeFileSync(tempPath, data);
  }
  replaceOutputPath(tempPath, finalPath, "file");
}

function hasUriScheme(value: string): boolean {
  return /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value);
}

function pathSegments(value: string): string[] {
  return value.split("/").filter((segment) => segment.length > 0);
}

export function isExternalRelationshipTarget(target: string | undefined, targetMode?: string): boolean {
  const mode = String(targetMode || "").toLowerCase();
  const value = String(target || "").trim();
  return mode === "external" || hasUriScheme(value);
}

export function normalizePackageEntryPath(name: string): string {
  const normalizedSlashes = String(name || "").replace(/\\/g, "/");
  if (!normalizedSlashes || normalizedSlashes.includes("\0") || normalizedSlashes.startsWith("/") || hasUriScheme(normalizedSlashes)) {
    throw new Error(`Unsafe Office package path: ${name}`);
  }
  if (pathSegments(normalizedSlashes).some((segment) => segment === "." || segment === "..")) {
    throw new Error(`Unsafe Office package path: ${name}`);
  }
  const normalized = path.posix.normalize(normalizedSlashes);
  if (!normalized || normalized === "." || normalized.startsWith("../") || normalized === ".." || path.posix.isAbsolute(normalized)) {
    throw new Error(`Unsafe Office package path: ${name}`);
  }
  return normalized;
}

function assertPackagePathInRoot(name: string, allowedRoot: string | undefined): void {
  if (!allowedRoot) {
    return;
  }
  const root = normalizePackageEntryPath(allowedRoot).replace(/\/?$/, "/");
  if (name !== root.slice(0, -1) && !name.startsWith(root)) {
    throw new Error(`Office package path escapes expected root: ${name}`);
  }
}

export function resolvePackagePath(baseFile: string, target: string, allowedRoot?: string): string {
  const base = normalizePackageEntryPath(baseFile);
  const normalizedTarget = String(target || "").replace(/\\/g, "/");
  if (!normalizedTarget || normalizedTarget.includes("\0") || hasUriScheme(normalizedTarget)) {
    throw new Error(`Unsafe Office relationship target: ${target}`);
  }

  const joined = normalizedTarget.startsWith("/")
    ? normalizedTarget.replace(/^\/+/, "")
    : path.posix.join(path.posix.dirname(base), normalizedTarget);
  const resolved = path.posix.normalize(joined);
  if (!resolved || resolved === "." || resolved.startsWith("../") || resolved === ".." || path.posix.isAbsolute(resolved)) {
    throw new Error(`Office relationship target escapes package root: ${target}`);
  }
  assertPackagePathInRoot(resolved, allowedRoot);
  return resolved;
}

function entryData(entry: AdmZip.IZipEntry, maxBytes: number): Buffer {
  if (Number(entry.header.size || 0) > maxBytes) {
    throw new Error(`Office package entry exceeds size limit: ${entry.entryName}`);
  }
  const data = entry.getData();
  if (data.length > maxBytes) {
    throw new Error(`Office package entry exceeds size limit: ${entry.entryName}`);
  }
  return data;
}

export function readXml(zip: AdmZip, name: string): XmlNode {
  const entryName = normalizePackageEntryPath(name);
  const entry = zip.getEntry(entryName);
  if (!entry) {
    throw new Error(`Office package entry not found: ${entryName}`);
  }
  return parseXml(entryData(entry, MAX_XML_ENTRY_BYTES).toString("utf8"));
}

export function readText(zip: AdmZip, name: string): string {
  const entryName = normalizePackageEntryPath(name);
  const entry = zip.getEntry(entryName);
  if (!entry) {
    throw new Error(`Office package entry not found: ${entryName}`);
  }
  return entryData(entry, MAX_XML_ENTRY_BYTES).toString("utf8");
}

export function readBuffer(zip: AdmZip, name: string, maxBytes = MAX_ENTRY_BYTES): Buffer {
  const entryName = normalizePackageEntryPath(name);
  const entry = zip.getEntry(entryName);
  if (!entry) {
    throw new Error(`Office package entry not found: ${entryName}`);
  }
  return entryData(entry, maxBytes);
}

export function hasEntry(zip: AdmZip, name: string): boolean {
  try {
    return Boolean(zip.getEntry(normalizePackageEntryPath(name)));
  } catch {
    return false;
  }
}

export function entries(zip: AdmZip): string[] {
  return zip.getEntries().map((entry: AdmZip.IZipEntry) => normalizePackageEntryPath(entry.entryName));
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
  const attributeKeys = new Set([
    "val",
    "space",
    "id",
    "embed",
    "link",
    "name",
    "descr",
    "cx",
    "cy",
    "Target",
    "TargetMode",
    "Type",
    "Id",
    "idx",
    "uri",
    "styleId",
    "styleID",
    "ilvl",
    "abstractNumId",
    "numId",
    "ascii",
    "eastAsia",
    "hAnsi",
    "cs",
    "hint",
    "type",
    "rsidR",
    "rsidRPr",
    "rsidRDefault",
    "rsidP",
    "paraId",
    "textId"
  ]);
  return Object.entries(node).filter(([key]) => !key.startsWith("#") && key !== ":@" && !attributeKeys.has(key));
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

export function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function quoteAttr(value: string): string {
  return `"${xmlEscape(value).replace(/"/g, "&quot;")}"`;
}

export function imageMimeType(packagePath: string): string | undefined {
  return IMAGE_MIME_BY_EXT[path.posix.extname(packagePath).toLowerCase()];
}

export function readImage(zip: AdmZip, packagePath: string): { data: Buffer; mime: string } | undefined {
  const entryName = normalizePackageEntryPath(packagePath);
  const mime = imageMimeType(entryName);
  if (!mime) {
    return undefined;
  }
  return {
    data: readBuffer(zip, entryName, MAX_IMAGE_BYTES),
    mime
  };
}

export function sanitizeMermaidText(value: string, fallback = "message"): string {
  const sanitized = String(value || "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/[`$]/g, "'")
    .replace(/[%{}[\]<>|]/g, " ")
    .replace(/-{2,}|={2,}|-{1,}>|<-{1,}/g, " ")
    .replace(/:/g, "：")
    .replace(/;/g, "；")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
  return sanitized || fallback;
}

export function round(value: number, digits = 2): number {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}
