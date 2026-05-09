#!/usr/bin/env python3
"""POC08: Word .docx を Draw.io XML に変換する。

Word は Markdown 向けに構造化する方針だが、比較用に Draw.io 化も試す。
画像は Draw.io で単体表示できるよう base64 data URI として埋め込む。
"""

from __future__ import annotations

import argparse
import base64
import json
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

EMU_PER_PIXEL = 9525
PAGE_X = 40
PAGE_Y = 40
TEXT_WIDTH = 620


def qname(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def read_xml(package: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(package.read(name))


def resolve_docx_path(path: Path) -> Path:
    if path.exists():
        return path
    candidates = sorted(path.parent.glob("*.docx"))
    normalized_name = path.name.replace("ペ", "ペ").replace("プ", "プ").replace("ポ", "ポ")
    for candidate in candidates:
        if candidate.name == normalized_name or path.name in candidate.name or "スペアミント" in path.name and "スペアミント" in candidate.name:
            return candidate
    raise FileNotFoundError(path)


def read_relationships(package: zipfile.ZipFile) -> dict[str, dict[str, str]]:
    root = read_xml(package, "word/_rels/document.xml.rels")
    rels: dict[str, dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        target = rel.attrib.get("Target", "")
        if target and not target.startswith("/"):
            target = f"word/{target}"
        rels[rel.attrib["Id"]] = {
            "type": rel.attrib.get("Type", ""),
            "target": target.lstrip("/"),
        }
    return rels


def paragraph_style(paragraph: ET.Element) -> str:
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return style.attrib.get(qname("w", "val"), "Normal") if style is not None else "Normal"


def direct_text(paragraph: ET.Element) -> str:
    parts: list[str] = []
    for run in paragraph.findall("w:r", NS):
        for node in run:
            name = node.tag.rsplit("}", 1)[-1]
            if name == "t":
                parts.append(node.text or "")
            elif name in {"br", "cr"}:
                parts.append("\n")
            elif name == "tab":
                parts.append("\t")
    return "".join(parts).strip()


def all_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def data_uri(package: zipfile.ZipFile, image_path: str) -> str:
    suffix = Path(image_path).suffix.lower().lstrip(".")
    mime = "jpeg" if suffix in {"jpg", "jpeg"} else suffix or "octet-stream"
    encoded = base64.b64encode(package.read(image_path)).decode("ascii")
    # Draw.io style は ; 区切りなので、;base64 を避けた data URI にする。
    return f"data:image/{mime},{encoded}"


def parse_drawing(drawing: ET.Element, package: zipfile.ZipFile, rels: dict[str, dict[str, str]]) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    doc_pr = drawing.find(".//wp:docPr", NS)
    extent = drawing.find(".//wp:extent", NS)
    base = {
        "name": doc_pr.attrib.get("name") if doc_pr is not None else "",
        "description": doc_pr.attrib.get("descr") if doc_pr is not None else "",
    }
    if extent is not None:
        base["width"] = round(int(extent.attrib.get("cx", "0")) / EMU_PER_PIXEL, 2)
        base["height"] = round(int(extent.attrib.get("cy", "0")) / EMU_PER_PIXEL, 2)

    for blip in drawing.findall(".//a:blip", NS):
        rid = blip.attrib.get(qname("r", "embed"), "")
        target = rels.get(rid, {}).get("target", "")
        if not target:
            continue
        items.append({
            **base,
            "type": "image",
            "path": target,
            "data_uri": data_uri(package, target),
        })

    shape_text = all_text(drawing)
    if shape_text and not items:
        labels = re.findall(r"オブジェクト[^オブジェクト\s]+", shape_text)
        items.append({
            **base,
            "type": "shape_text",
            "text": shape_text,
            "labels": labels or [shape_text],
        })
    return items


def parse_docx(source: Path) -> dict[str, Any]:
    source = resolve_docx_path(source)
    with zipfile.ZipFile(source) as package:
        rels = read_relationships(package)
        document = read_xml(package, "word/document.xml")
        body = document.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml に body がありません")

        blocks: list[dict[str, Any]] = []
        for index, child in enumerate(body):
            if child.tag != qname("w", "p"):
                continue
            drawings: list[dict[str, Any]] = []
            for drawing in child.findall(".//w:drawing", NS):
                drawings.extend(parse_drawing(drawing, package, rels))
            blocks.append({
                "index": index,
                "style": paragraph_style(child),
                "text": direct_text(child),
                "drawings": drawings,
            })
    return {
        "document": {
            "source": source.as_posix(),
            "format": "docx",
            "output": "drawio",
            "image_output": "base64",
        },
        "blocks": blocks,
    }


def text_style(style: str) -> str:
    if style == "Title":
        return "text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;fontSize=28;fontStyle=1;"
    if style == "Subtitle":
        return "text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;fontSize=16;fontStyle=2;"
    if style.startswith("Heading"):
        return "text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;fontSize=20;fontStyle=1;"
    return "text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=top;fontSize=12;"


def add_text_cell(cells: list[str], cell_id: str, value: str, style: str, x: float, y: float, width: float, height: float) -> None:
    cells.append(f'        <mxCell id="{cell_id}" value={quoteattr(value)} style={quoteattr(style)} vertex="1" parent="1">')
    cells.append(f'          <mxGeometry x="{x}" y="{y}" width="{width}" height="{height}" as="geometry"/>')
    cells.append("        </mxCell>")


def add_image_cell(cells: list[str], cell_id: str, image: dict[str, Any], x: float, y: float) -> float:
    width = min(float(image.get("width") or 240), 260)
    height = min(float(image.get("height") or 180), 360)
    style = f"shape=image;html=1;imageAspect=0;aspect=fixed;image={image.get('data_uri', '')};"
    cells.append(f'        <mxCell id="{cell_id}" value="" style={quoteattr(style)} vertex="1" parent="1">')
    cells.append(f'          <mxGeometry x="{x}" y="{y}" width="{width}" height="{height}" as="geometry"/>')
    cells.append("        </mxCell>")
    return height


def is_main_layout_image(image: dict[str, Any]) -> bool:
    path = str(image.get("path", ""))
    width = float(image.get("width") or 0)
    height = float(image.get("height") or 0)
    return not path.endswith("image3.png") and width > 100 and height > 100


def is_drawio_visible_image(image: dict[str, Any]) -> bool:
    return not str(image.get("path", "")).endswith("image3.png")


def make_drawio(intermediate: dict[str, Any]) -> str:
    cells = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mxfile host="app.diagrams.net">',
        '  <diagram id="POC08-WORD-1" name="Word to Draw.io">',
        '    <mxGraphModel dx="1000" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169">',
        "      <root>",
        '        <mxCell id="0"/>',
        '        <mxCell id="1" parent="0"/>',
    ]
    y = PAGE_Y
    pending_images: list[dict[str, Any]] = []
    shape_labels: list[str] = []

    for block in intermediate["blocks"]:
        images = [
            d for d in block["drawings"]
            if d["type"] == "image" and is_drawio_visible_image(d)
        ]
        shapes = [d for d in block["drawings"] if d["type"] == "shape_text"]
        if shapes:
            for shape in shapes:
                shape_labels.extend(shape.get("labels") or [])
            continue
        if images and not block["text"]:
            main_images = [image for image in images if is_main_layout_image(image)]
            small_images = [image for image in images if not is_main_layout_image(image)]
            pending_images.extend(main_images)
            for image in small_images:
                y += add_image_cell(cells, f"image-{block['index']}-{len(small_images)}", image, PAGE_X, y) + 16
            continue

        if pending_images and block["text"]:
            image = pending_images.pop(0)
            used_height = add_image_cell(cells, f"image-{block['index']}", image, PAGE_X, y)
            add_text_cell(cells, f"text-{block['index']}", block["text"], text_style(block["style"]), PAGE_X + 300, y, 360, max(80, used_height))
            y += max(used_height, 110) + 24
        elif block["text"]:
            height = 36 if block["style"] in {"Title", "Subtitle"} or block["style"].startswith("Heading") else max(36, 18 * (block["text"].count("\n") + 2))
            add_text_cell(cells, f"text-{block['index']}", block["text"], text_style(block["style"]), PAGE_X, y, TEXT_WIDTH, height)
            y += height + 12

        while pending_images:
            image = pending_images.pop(0)
            y += add_image_cell(cells, f"image-{block['index']}-{len(pending_images)}", image, PAGE_X, y) + 16

    if len(shape_labels) >= 2:
        y += 20
        add_text_cell(cells, "shape-a", shape_labels[0], "rounded=1;arcSize=12;whiteSpace=wrap;html=1;fillColor=#CFE2F3;strokeColor=#000000;align=center;verticalAlign=middle;", PAGE_X, y, 180, 70)
        add_text_cell(cells, "shape-b", shape_labels[1], "rounded=1;arcSize=12;whiteSpace=wrap;html=1;fillColor=#F4CCCC;strokeColor=#000000;align=center;verticalAlign=middle;", PAGE_X + 300, y, 180, 70)
        cells.append('        <mxCell id="edge-a-b" value="" style="edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;endArrow=classic;strokeColor=#000000;" edge="1" parent="1" source="shape-a" target="shape-b">')
        cells.append('          <mxGeometry relative="1" as="geometry"/>')
        cells.append("        </mxCell>")

    cells.extend([
        "      </root>",
        "    </mxGraphModel>",
        "  </diagram>",
        "</mxfile>",
    ])
    return "\n".join(cells)


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parents[4]
    default_input = repo_root / "doc" / "30.test" / "00.pocTestData" / "ペット プロフィール - スペアミント (1).docx"
    parser = argparse.ArgumentParser(description="POC08: docx to Draw.io")
    parser.add_argument("--input", type=Path, default=default_input)
    parser.add_argument("--json", type=Path, default=script_dir / "poc08_word_intermediate.json")
    parser.add_argument("--drawio", type=Path, default=script_dir / "poc08_word_output.dio.xml")
    args = parser.parse_args()

    intermediate = parse_docx(args.input)
    args.json.write_text(json.dumps(intermediate, ensure_ascii=False, indent=2), encoding="utf-8")
    args.drawio.write_text(make_drawio(intermediate), encoding="utf-8")
    print(f"Wrote {args.json}")
    print(f"Wrote {args.drawio}")


if __name__ == "__main__":
    main()
