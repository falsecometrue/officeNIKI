#!/usr/bin/env python3
"""POC01: Word .docx を中間 JSON に解析し、Markdown を生成する。

方針:
- .docx は zip なので、word/document.xml と relationships を直接読む。
- 見出し/段落/表/画像を JSON に集約してから Markdown を生成する。
- Word 図形は Markdown で完全再現しづらいため、画像フォールバック候補と
  Mermaid 化候補を JSON に残す。
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import struct
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

EMU_PER_PIXEL = 9525


@dataclass
class Context:
    source_docx: Path
    output_dir: Path
    resource_dir: Path
    rels: dict[str, dict[str, str]]
    copied_resources: dict[str, str]


def qname(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def read_xml(package: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(package.read(name))


def parse_relationships(package: zipfile.ZipFile) -> dict[str, dict[str, str]]:
    """document.xml から参照される画像などの rId を解決する。"""
    rels_root = read_xml(package, "word/_rels/document.xml.rels")
    rels: dict[str, dict[str, str]] = {}
    for rel in rels_root.findall("rel:Relationship", NS):
        rid = rel.attrib["Id"]
        target = rel.attrib["Target"]
        if not target.startswith("/"):
            target = f"word/{target}"
        rels[rid] = {
            "type": rel.attrib.get("Type", ""),
            "target": target,
        }
    return rels


def parse_image_size(data: bytes, suffix: str) -> dict[str, int] | None:
    """PNG/JPEG のピクセルサイズだけを軽量に読む。"""
    suffix = suffix.lower()
    if suffix == ".png" and data.startswith(b"\x89PNG\r\n\x1a\n"):
        width, height = struct.unpack(">II", data[16:24])
        return {"width_px": width, "height_px": height}

    if suffix in {".jpg", ".jpeg"} and data.startswith(b"\xff\xd8"):
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            length = struct.unpack(">H", data[i + 2 : i + 4])[0]
            if 0xC0 <= marker <= 0xC3:
                height, width = struct.unpack(">HH", data[i + 5 : i + 9])
                return {"width_px": width, "height_px": height}
            i += 2 + length
    return None


def copy_resource(ctx: Context, package: zipfile.ZipFile, target: str) -> dict[str, Any]:
    """画像リソースを resources 配下へコピーし、Markdown 参照パスを返す。"""
    if target in ctx.copied_resources:
        rel_path = ctx.copied_resources[target]
    else:
        data = package.read(target)
        dest = ctx.resource_dir / Path(target).name
        # 同名ファイル衝突を避ける。通常 word/media は一意だが POC として保険を入れる。
        if dest.exists():
            stem = dest.stem
            suffix = dest.suffix
            n = 2
            while (ctx.resource_dir / f"{stem}_{n}{suffix}").exists():
                n += 1
            dest = ctx.resource_dir / f"{stem}_{n}{suffix}"
        dest.write_bytes(data)
        rel_path = dest.relative_to(ctx.output_dir).as_posix()
        ctx.copied_resources[target] = rel_path

    data = package.read(target)
    size = parse_image_size(data, Path(target).suffix) or {}
    return {
        "source": target,
        "path": rel_path,
        "content_type_hint": Path(target).suffix.lower().lstrip("."),
        **size,
    }


def paragraph_style(paragraph: ET.Element) -> str:
    pstyle = paragraph.find("w:pPr/w:pStyle", NS)
    return pstyle.attrib.get(qname("w", "val"), "Normal") if pstyle is not None else "Normal"


def direct_paragraph_text(paragraph: ET.Element) -> str:
    """段落直下のテキストを読む。テキストボックス内の文字は本文として重複させない。"""
    parts: list[str] = []
    for child in paragraph:
        if child.tag != qname("w", "r"):
            continue
        for run_child in child:
            name = local_name(run_child.tag)
            if name == "t":
                parts.append(run_child.text or "")
            elif name in {"tab"}:
                parts.append("\t")
            elif name in {"br", "cr"}:
                parts.append("\n")
    return "".join(parts).strip()


def all_text(element: ET.Element) -> str:
    return "".join(t.text or "" for t in element.findall(".//w:t", NS)).strip()


def extract_object_labels(text: str) -> list[str]:
    """サンプルの「オブジェクトAオブジェクトB」のような連結ラベルを分割する。"""
    labels = re.findall(r"オブジェクト[^オブジェクト\s]+", text)
    if len(labels) >= 2:
        return labels
    return [text] if text else []


def parse_drawing(drawing: ET.Element, ctx: Context, package: zipfile.ZipFile) -> list[dict[str, Any]]:
    """w:drawing を画像または図形候補として取り出す。"""
    doc_pr = drawing.find(".//wp:docPr", NS)
    extent = drawing.find(".//wp:extent", NS)
    base = {
        "doc_pr_id": doc_pr.attrib.get("id") if doc_pr is not None else None,
        "name": doc_pr.attrib.get("name") if doc_pr is not None else None,
        "description": doc_pr.attrib.get("descr") if doc_pr is not None else None,
    }
    if extent is not None:
        cx = int(extent.attrib.get("cx", "0"))
        cy = int(extent.attrib.get("cy", "0"))
        base["size"] = {
            "cx_emu": cx,
            "cy_emu": cy,
            "width_px": round(cx / EMU_PER_PIXEL, 2),
            "height_px": round(cy / EMU_PER_PIXEL, 2),
        }

    elements: list[dict[str, Any]] = []
    for blip in drawing.findall(".//a:blip", NS):
        rid = blip.attrib.get(qname("r", "embed"))
        rel = ctx.rels.get(rid or "")
        if not rid or not rel:
            continue
        resource = copy_resource(ctx, package, rel["target"])
        elements.append({
            **base,
            "type": "image",
            "relationship_id": rid,
            "resource": resource,
        })

    shape_text = all_text(drawing)
    if shape_text and not elements:
        labels = extract_object_labels(shape_text)
        mermaid = None
        if len(labels) >= 2:
            mermaid = {
                "type": "flowchart_lr",
                "code": "\n".join([
                    "flowchart LR",
                    f"  A[{labels[0]}] --> B[{labels[1]}]",
                ]),
                "note": "図形内テキストから推定した候補。矢印や配置は docx XML だけでは確定できない。",
            }
        elements.append({
            **base,
            "type": "shape_text",
            "text": shape_text,
            "labels": labels,
            "mermaid_candidate": mermaid,
        })

    return elements


def parse_table(table: ET.Element) -> dict[str, Any]:
    rows: list[list[str]] = []
    for tr in table.findall("w:tr", NS):
        row: list[str] = []
        for tc in tr.findall("w:tc", NS):
            row.append(all_text(tc))
        rows.append(row)
    return {"type": "table", "rows": rows}


def parse_docx(source_docx: Path, output_dir: Path) -> dict[str, Any]:
    resource_dir = output_dir / "resources"
    if resource_dir.exists():
        shutil.rmtree(resource_dir)
    resource_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source_docx) as package:
        ctx = Context(
            source_docx=source_docx,
            output_dir=output_dir,
            resource_dir=resource_dir,
            rels=parse_relationships(package),
            copied_resources={},
        )
        document = read_xml(package, "word/document.xml")
        body = document.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml に body がありません")

        blocks: list[dict[str, Any]] = []
        images: list[dict[str, Any]] = []
        shapes: list[dict[str, Any]] = []

        for index, child in enumerate(body):
            if child.tag == qname("w", "p"):
                drawings: list[dict[str, Any]] = []
                for drawing in child.findall(".//w:drawing", NS):
                    drawings.extend(parse_drawing(drawing, ctx, package))
                for drawing in drawings:
                    if drawing["type"] == "image":
                        images.append(drawing)
                    else:
                        shapes.append(drawing)

                blocks.append({
                    "type": "paragraph",
                    "index": index,
                    "style": paragraph_style(child),
                    "text": direct_paragraph_text(child),
                    "drawings": drawings,
                })
            elif child.tag == qname("w", "tbl"):
                table = parse_table(child)
                table["index"] = index
                blocks.append(table)

    return {
        "document": {
            "source": source_docx.as_posix(),
            "format": "docx",
        },
        "blocks": blocks,
        "images": images,
        "shapes": shapes,
        "conversion_notes": [
            "段落/見出し/画像/表は Markdown へ直接変換する。",
            "Word 図形は DrawingML の座標や形状情報が Markdown と対応しづらいため、画像化または Mermaid 化を検討対象にする。",
            "本 POC は外部ライブラリを使わず、docx の zip/OpenXML を直接解析している。",
        ],
    }


def md_escape_table_cell(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", "<br>")


def markdown_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    normalized = [row + [""] * (width - len(row)) for row in rows]
    header = "| " + " | ".join(md_escape_table_cell(v) for v in normalized[0]) + " |"
    sep = "| " + " | ".join("---" for _ in range(width)) + " |"
    body = [
        "| " + " | ".join(md_escape_table_cell(v) for v in row) + " |"
        for row in normalized[1:]
    ]
    return "\n".join([header, sep, *body])


def markdown_for_block(block: dict[str, Any]) -> list[str]:
    if block["type"] == "table":
        table_md = markdown_table(block["rows"])
        return [table_md] if table_md else []

    text = block.get("text", "")
    style = block.get("style", "Normal")
    lines: list[str] = []

    if text:
        if style == "Title":
            lines.append(f"# {text}")
        elif style == "Subtitle":
            lines.append(f"*{text}*")
        elif style.startswith("Heading"):
            level = int(re.sub(r"[^0-9]", "", style) or "1") + 1
            level = min(max(level, 2), 6)
            lines.append(f"{'#' * level} {text}")
        else:
            lines.append(text)

    for drawing in block.get("drawings", []):
        if drawing["type"] == "image":
            resource = drawing["resource"]
            if resource.get("width_px", 0) <= 1 and resource.get("height_px", 0) <= 1:
                continue
            alt = drawing.get("description") or drawing.get("name") or Path(resource["path"]).name
            lines.append(f"![{alt}]({resource['path']})")
        elif drawing["type"] == "shape_text":
            mermaid = drawing.get("mermaid_candidate")
            if mermaid:
                lines.extend(["```mermaid", mermaid["code"], "```"])
            else:
                lines.append(f"> 図形テキスト: {drawing.get('text', '')}")

    return lines


def build_markdown(intermediate: dict[str, Any]) -> str:
    chunks: list[str] = []
    for block in intermediate["blocks"]:
        lines = markdown_for_block(block)
        if lines:
            chunks.append("\n".join(lines))
    return "\n\n".join(chunks).rstrip() + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="POC01 docx to intermediate JSON and Markdown")
    parser.add_argument("source_docx", type=Path)
    parser.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--json", type=str, default="poc01_intermediate.json")
    parser.add_argument("--md", type=str, default="poc01_output.md")
    args = parser.parse_args()

    out_dir = args.out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    intermediate = parse_docx(args.source_docx.resolve(), out_dir)

    (out_dir / args.json).write_text(
        json.dumps(intermediate, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (out_dir / args.md).write_text(build_markdown(intermediate), encoding="utf-8")

    print(f"JSON: {(out_dir / args.json).as_posix()}")
    print(f"Markdown: {(out_dir / args.md).as_posix()}")
    print(f"Resources: {(out_dir / 'resources').as_posix()}")


if __name__ == "__main__":
    main()
