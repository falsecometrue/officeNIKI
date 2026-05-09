#!/usr/bin/env python3
import argparse
import base64
import json
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

EMU_PER_INCH = 914400
PX_PER_INCH = 96
DEFAULT_COL_WIDTH = 12.63
DEFAULT_ROW_HEIGHT_PT = 15.75
TABLE_ORIGIN_X = 40
TABLE_ORIGIN_Y = 280


# Excel DrawingML uses EMU. Draw.io geometry is easier to handle as px.
def emu_to_px(value: object) -> float:
    return round(int(value) / EMU_PER_INCH * PX_PER_INCH, 2)


# Excel column names are base-26 numbers: A=1, B=2, ..., AA=27.
def col_name_to_index(col_name: str) -> int:
    index = 0
    for ch in col_name:
        index = index * 26 + ord(ch.upper()) - ord("A") + 1
    return index


def split_cell_ref(ref: str) -> Tuple[int, int]:
    match = re.match(r"([A-Z]+)([0-9]+)", ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    return col_name_to_index(match.group(1)), int(match.group(2))


def read_xml(package: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(package.read(name))


# Relationship files connect workbook -> sheet and sheet -> drawing/media.
def read_relationships(package: zipfile.ZipFile, name: str) -> Dict[str, Dict[str, str]]:
    if name not in package.namelist():
        return {}
    root = read_xml(package, name)
    relationships = {}
    for rel in root.findall("pkgrel:Relationship", NS):
        relationships[rel.attrib["Id"]] = {
            "type": rel.attrib.get("Type", ""),
            "target": rel.attrib.get("Target", ""),
        }
    return relationships


def resolve_package_path(base_file: str, target: str) -> str:
    base_dir = posixpath.dirname(base_file)
    return posixpath.normpath(posixpath.join(base_dir, target))


# sharedStrings.xml stores most Excel text once, while sheet cells keep indexes.
def read_shared_strings(package: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in package.namelist():
        return []
    root = read_xml(package, "xl/sharedStrings.xml")
    values = []
    for si in root.findall("main:si", NS):
        text = "".join(t.text or "" for t in si.findall(".//main:t", NS))
        values.append(text)
    return values


def parse_workbook(package: zipfile.ZipFile) -> List[Dict[str, str]]:
    workbook = read_xml(package, "xl/workbook.xml")
    workbook_rels = read_relationships(package, "xl/_rels/workbook.xml.rels")
    sheets = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        rel_id = sheet.attrib.get(f"{{{NS['r']}}}id")
        target = workbook_rels.get(rel_id, {}).get("target", "")
        path = resolve_package_path("xl/workbook.xml", target) if target else ""
        sheets.append({
            "name": sheet.attrib.get("name", ""),
            "sheet_id": sheet.attrib.get("sheetId", ""),
            "rel_id": rel_id or "",
            "path": path,
        })
    return sheets


# styles.xml cellXfs are referenced from sheet cells by style_id.
def parse_styles(package: zipfile.ZipFile) -> List[Dict[str, object]]:
    if "xl/styles.xml" not in package.namelist():
        return []
    root = read_xml(package, "xl/styles.xml")
    styles = []
    for index, xf in enumerate(root.findall("main:cellXfs/main:xf", NS)):
        styles.append({
            "style_id": index,
            "font_id": int(xf.attrib.get("fontId", "0")),
            "fill_id": int(xf.attrib.get("fillId", "0")),
            "border_id": int(xf.attrib.get("borderId", "0")),
            "num_fmt_id": int(xf.attrib.get("numFmtId", "0")),
            "apply_border": xf.attrib.get("applyBorder") == "1",
            "apply_alignment": xf.attrib.get("applyAlignment") == "1",
        })
    return styles


# Extract only cells that Excel wrote into sheetData. Empty unstyled cells are not generated.
def parse_cells(sheet_root: ET.Element, shared_strings: List[str]) -> Tuple[List[Dict[str, object]], List[Dict[str, object]], List[Dict[str, object]]]:
    rows = []
    cells = []
    cols = []
    for col in sheet_root.findall("main:cols/main:col", NS):
        cols.append({
            "min": int(col.attrib.get("min", "0")),
            "max": int(col.attrib.get("max", "0")),
            "width": float(col.attrib.get("width", "0")),
            "custom_width": col.attrib.get("customWidth") == "1",
        })
    for row in sheet_root.findall("main:sheetData/main:row", NS):
        row_index = int(row.attrib["r"])
        rows.append({
            "index": row_index,
            "height": float(row.attrib["ht"]) if "ht" in row.attrib else None,
        })
        for cell in row.findall("main:c", NS):
            ref = cell.attrib["r"]
            col_index, row_number = split_cell_ref(ref)
            raw_value = cell.findtext("main:v", default="", namespaces=NS)
            cell_type = cell.attrib.get("t", "n")
            value = raw_value
            if cell_type == "s" and raw_value:
                value = shared_strings[int(raw_value)]
            cells.append({
                "ref": ref,
                "row": row_number,
                "col": col_index,
                "style_id": int(cell.attrib.get("s", "0")),
                "type": cell_type,
                "raw_value": raw_value,
                "value": value,
            })
    return rows, cols, cells


def parse_merged_cells(sheet_root: ET.Element) -> List[Dict[str, str]]:
    merged = []
    for merge_cell in sheet_root.findall("main:mergeCells/main:mergeCell", NS):
        ref = merge_cell.attrib.get("ref", "")
        if ":" in ref:
            start, end = ref.split(":", 1)
            merged.append({"ref": ref, "start": start, "end": end})
    return merged


def text_from_drawing(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//a:t", NS))


def find_color(node: ET.Element, path: str) -> Optional[str]:
    color = node.find(path, NS)
    if color is None:
        return None
    return f"#{color.attrib['val']}"


# Anchors tell where a drawing is attached to the sheet grid.
def parse_anchor(anchor: ET.Element) -> Dict[str, object]:
    from_node = anchor.find("xdr:from", NS)
    ext = anchor.find("xdr:ext", NS)
    result: Dict[str, object] = {
        "type": anchor.tag.rsplit("}", 1)[-1],
        "from": None,
        "ext": None,
    }
    if from_node is not None:
        result["from"] = {
            "col": int(from_node.findtext("xdr:col", default="0", namespaces=NS)),
            "row": int(from_node.findtext("xdr:row", default="0", namespaces=NS)),
            "colOff_emu": int(from_node.findtext("xdr:colOff", default="0", namespaces=NS)),
            "rowOff_emu": int(from_node.findtext("xdr:rowOff", default="0", namespaces=NS)),
        }
    if ext is not None:
        result["ext"] = {
            "cx_emu": int(ext.attrib.get("cx", "0")),
            "cy_emu": int(ext.attrib.get("cy", "0")),
            "width": emu_to_px(ext.attrib.get("cx", "0")),
            "height": emu_to_px(ext.attrib.get("cy", "0")),
        }
    return result


# Convert one DrawingML shape/connector into the intermediate drawing object.
def parse_drawing_shape(node: ET.Element) -> Dict[str, object]:
    nv = node.find(".//xdr:cNvPr", NS)
    xfrm = node.find(".//a:xfrm", NS)
    geom = node.find(".//a:prstGeom", NS)
    off = xfrm.find("a:off", NS) if xfrm is not None else None
    ext = xfrm.find("a:ext", NS) if xfrm is not None else None
    tail = node.find(".//a:ln/a:tailEnd", NS)
    head = node.find(".//a:ln/a:headEnd", NS)
    start_cxn = node.find(".//a:stCxn", NS)
    end_cxn = node.find(".//a:endCxn", NS)
    kind = "edge" if node.tag.endswith("cxnSp") else "vertex"
    shape: Dict[str, object] = {
        "id": nv.attrib.get("id") if nv is not None else "",
        "name": nv.attrib.get("name") if nv is not None else "",
        "kind": kind,
        "preset": geom.attrib.get("prst") if geom is not None else "",
        "text": text_from_drawing(node),
        "fill": find_color(node, ".//a:solidFill/a:srgbClr"),
        "stroke": find_color(node, ".//a:ln/a:solidFill/a:srgbClr"),
        "headEnd": head.attrib.get("type") if head is not None else None,
        "tailEnd": tail.attrib.get("type") if tail is not None else None,
        "startConnectionId": start_cxn.attrib.get("id") if start_cxn is not None else None,
        "endConnectionId": end_cxn.attrib.get("id") if end_cxn is not None else None,
    }
    if off is not None:
        shape.update({
            "x_emu": int(off.attrib.get("x", "0")),
            "y_emu": int(off.attrib.get("y", "0")),
            "x": emu_to_px(off.attrib.get("x", "0")),
            "y": emu_to_px(off.attrib.get("y", "0")),
        })
    if ext is not None:
        shape.update({
            "width_emu": int(ext.attrib.get("cx", "0")),
            "height_emu": int(ext.attrib.get("cy", "0")),
            "width": emu_to_px(ext.attrib.get("cx", "0")),
            "height": emu_to_px(ext.attrib.get("cy", "0")),
        })
    return shape


# drawing*.xml can contain anchors, grouped shapes, connectors, and pictures.
def parse_drawings(package: zipfile.ZipFile, drawing_path: str) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    if drawing_path not in package.namelist():
        return [], []
    root = read_xml(package, drawing_path)
    anchors = []
    drawings = []
    for anchor in list(root):
        if not anchor.tag.endswith(("oneCellAnchor", "twoCellAnchor", "absoluteAnchor")):
            continue
        anchors.append(parse_anchor(anchor))
        containers = [anchor]
        containers.extend(anchor.findall("xdr:grpSp", NS))
        for container in containers:
            for child in list(container):
                if child.tag.endswith("sp") or child.tag.endswith("cxnSp") or child.tag.endswith("pic"):
                    if child.tag.endswith("pic"):
                        continue
                    drawings.append(parse_drawing_shape(child))
    infer_edge_sources(drawings)
    return drawings, anchors


# Some connectors have only one explicit connection. Fill the missing side from nearby vertices.
def infer_edge_sources(drawings: List[Dict[str, object]]) -> None:
    vertices = [d for d in drawings if d.get("kind") == "vertex"]
    for edge in [d for d in drawings if d.get("kind") == "edge"]:
        if edge.get("startConnectionId"):
            edge["sourceId"] = edge["startConnectionId"]
        else:
            edge["sourceId"] = nearest_vertex_id(edge, vertices, use_end=False)
        edge["targetId"] = edge.get("endConnectionId") or nearest_vertex_id(edge, vertices, use_end=True)


def nearest_vertex_id(edge: Dict[str, object], vertices: List[Dict[str, object]], use_end: bool) -> Optional[str]:
    ex = float(edge.get("x", 0))
    ey = float(edge.get("y", 0))
    if use_end:
        ex += float(edge.get("width", 0))
        ey += float(edge.get("height", 0))
    best_id = None
    best_distance = None
    for vertex in vertices:
        vx = float(vertex.get("x", 0)) + float(vertex.get("width", 0)) / 2
        vy = float(vertex.get("y", 0)) + float(vertex.get("height", 0)) / 2
        distance = (vx - ex) ** 2 + (vy - ey) ** 2
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_id = str(vertex.get("id"))
    return best_id


# Keep embedded media as base64 so the intermediate JSON can become a single draw.io file later.
def read_images(package: zipfile.ZipFile) -> List[Dict[str, object]]:
    images = []
    for name in package.namelist():
        if not name.startswith("xl/media/"):
            continue
        suffix = Path(name).suffix.lower().lstrip(".")
        mime = "jpeg" if suffix in {"jpg", "jpeg"} else suffix or "octet-stream"
        data = base64.b64encode(package.read(name)).decode("ascii")
        images.append({
            "path": name,
            "mime": f"image/{mime}",
            "base64": data,
        })
    return images


# Main xlsx -> intermediate JSON pipeline.
def build_intermediate(xlsx_path: Path) -> Dict[str, object]:
    with zipfile.ZipFile(xlsx_path) as package:
        shared_strings = read_shared_strings(package)
        styles = parse_styles(package)
        sheet_defs = parse_workbook(package)
        images = read_images(package)
        sheets = []
        for sheet_def in sheet_defs:
            sheet_root = read_xml(package, sheet_def["path"])
            rows, cols, cells = parse_cells(sheet_root, shared_strings)
            merged_cells = parse_merged_cells(sheet_root)
            rel_path = f"{posixpath.dirname(sheet_def['path'])}/_rels/{posixpath.basename(sheet_def['path'])}.rels"
            sheet_rels = read_relationships(package, rel_path)
            drawings = []
            anchors = []
            for rel in sheet_rels.values():
                if rel["type"].endswith("/drawing"):
                    drawing_path = resolve_package_path(sheet_def["path"], rel["target"])
                    drawings, anchors = parse_drawings(package, drawing_path)
            sheets.append({
                "name": sheet_def["name"],
                "path": sheet_def["path"],
                "rows": rows,
                "cols": cols,
                "cells": cells,
                "styles": styles,
                "merged_cells": merged_cells,
                "drawings": drawings,
                "images": images,
                "anchors": anchors,
            })
        return {
            "workbook": {
                "source": str(xlsx_path),
                "format": "xlsx-zip",
                "sheets_count": len(sheets),
            },
            "sheets": sheets,
        }


def drawio_shape_style(shape: Dict[str, object]) -> str:
    fill = shape.get("fill") or "#FFFFFF"
    stroke = shape.get("stroke") or "#000000"
    return f"rounded=0;whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};"


def excel_col_width_to_px(width: float) -> float:
    # Approximation used for layout POC. Excel's exact width depends on default font metrics.
    return round(width * 7 + 5, 2)


def excel_row_height_to_px(height_pt: float) -> float:
    return round(height_pt * PX_PER_INCH / 72, 2)


def col_width_map(sheet: Dict[str, object]) -> Dict[int, float]:
    widths: Dict[int, float] = {}
    for col in sheet.get("cols", []):
        for index in range(int(col["min"]), int(col["max"]) + 1):
            widths[index] = excel_col_width_to_px(float(col.get("width") or DEFAULT_COL_WIDTH))
    return widths


def row_height_map(sheet: Dict[str, object]) -> Dict[int, float]:
    heights: Dict[int, float] = {}
    for row in sheet.get("rows", []):
        height = row.get("height") or DEFAULT_ROW_HEIGHT_PT
        heights[int(row["index"])] = excel_row_height_to_px(float(height))
    return heights


def cell_geometry(cell: Dict[str, object], col_widths: Dict[int, float], row_heights: Dict[int, float]) -> Dict[str, float]:
    col = int(cell["col"])
    row = int(cell["row"])
    default_col_px = excel_col_width_to_px(DEFAULT_COL_WIDTH)
    default_row_px = excel_row_height_to_px(DEFAULT_ROW_HEIGHT_PT)
    x = TABLE_ORIGIN_X + sum(col_widths.get(i, default_col_px) for i in range(1, col))
    y = TABLE_ORIGIN_Y + sum(row_heights.get(i, default_row_px) for i in range(1, row))
    return {
        "x": round(x, 2),
        "y": round(y, 2),
        "width": col_widths.get(col, default_col_px),
        "height": row_heights.get(row, default_row_px),
    }


def drawio_cell_style(cell: Dict[str, object], styles: Dict[int, Dict[str, object]]) -> str:
    style = styles.get(int(cell.get("style_id", 0)), {})
    stroke = "#000000" if style.get("apply_border") else "#D9D9D9"
    return (
        "rounded=0;whiteSpace=wrap;html=1;"
        "fillColor=#FFFFFF;"
        f"strokeColor={stroke};"
        "align=center;verticalAlign=middle;fontSize=11;"
    )


def append_cell_vertices(xml_cells: List[str], sheet: Dict[str, object]) -> None:
    styles = {int(style["style_id"]): style for style in sheet.get("styles", [])}
    col_widths = col_width_map(sheet)
    row_heights = row_height_map(sheet)
    for cell in sheet.get("cells", []):
        geometry = cell_geometry(cell, col_widths, row_heights)
        cell_id = f"cell-{cell['ref']}"
        value = quoteattr(str(cell.get("value", "")))
        style = quoteattr(drawio_cell_style(cell, styles))
        xml_cells.append(f'        <mxCell id="{cell_id}" value={value} style={style} vertex="1" parent="1">')
        xml_cells.append(f'          <mxGeometry x="{geometry["x"]}" y="{geometry["y"]}" width="{geometry["width"]}" height="{geometry["height"]}" as="geometry"/>')
        xml_cells.append("        </mxCell>")


def append_drawing_vertices(xml_cells: List[str], sheet: Dict[str, object]) -> None:
    for shape in sheet.get("drawings", []):
        if shape.get("kind") != "vertex":
            continue
        cell_id = f"shape-{shape['id']}"
        value = quoteattr(str(shape.get("text", "")))
        style = quoteattr(drawio_shape_style(shape))
        xml_cells.append(f'        <mxCell id="{cell_id}" value={value} style={style} vertex="1" parent="1">')
        xml_cells.append(f'          <mxGeometry x="{shape.get("x", 0)}" y="{shape.get("y", 0)}" width="{shape.get("width", 120)}" height="{shape.get("height", 60)}" as="geometry"/>')
        xml_cells.append("        </mxCell>")


def append_drawing_edges(xml_cells: List[str], sheet: Dict[str, object]) -> None:
    for shape in sheet.get("drawings", []):
        if shape.get("kind") != "edge":
            continue
        cell_id = f"edge-{shape['id']}"
        source = f"shape-{shape['sourceId']}" if shape.get("sourceId") else ""
        target = f"shape-{shape['targetId']}" if shape.get("targetId") else ""
        stroke = shape.get("stroke") or "#000000"
        style = quoteattr(f"edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;strokeColor={stroke};endArrow=classic;")
        source_attr = f' source="{source}"' if source else ""
        target_attr = f' target="{target}"' if target else ""
        xml_cells.append(f'        <mxCell id="{cell_id}" value="" style={style} edge="1" parent="1"{source_attr}{target_attr}>')
        xml_cells.append('          <mxGeometry relative="1" as="geometry"/>')
        xml_cells.append("        </mxCell>")


# Intermediate JSON -> Draw.io XML. This reflects extracted cells, shapes, and edges.
def make_drawio(intermediate: Dict[str, object]) -> str:
    sheet = intermediate["sheets"][0]
    cells = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mxfile host="app.diagrams.net">',
        '  <diagram id="POC06" name="Page-1">',
        '    <mxGraphModel dx="1000" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169">',
        '      <root>',
        '        <mxCell id="0"/>',
        '        <mxCell id="1" parent="0"/>',
    ]
    append_cell_vertices(cells, sheet)
    append_drawing_vertices(cells, sheet)
    append_drawing_edges(cells, sheet)
    cells.extend([
        "      </root>",
        "    </mxGraphModel>",
        "  </diagram>",
        "</mxfile>",
    ])
    return "\n".join(cells)


# CLI entrypoint for repeatable POC runs.
def main() -> None:
    script_dir = Path(__file__).resolve().parent
    default_input = script_dir.parents[2] / "30.test" / "testData" / "テストデータ.xlsx"
    parser = argparse.ArgumentParser(description="POC06: xlsx zip XML to intermediate JSON and Draw.io")
    parser.add_argument("--input", type=Path, default=default_input)
    parser.add_argument("--json", type=Path, default=script_dir / "poc06_intermediate.json")
    parser.add_argument("--drawio", type=Path, default=script_dir / "poc06_output.drawio")
    args = parser.parse_args()

    intermediate = build_intermediate(args.input)
    args.json.write_text(json.dumps(intermediate, ensure_ascii=False, indent=2), encoding="utf-8")
    args.drawio.write_text(make_drawio(intermediate), encoding="utf-8")
    print(f"Wrote {args.json}")
    print(f"Wrote {args.drawio}")


if __name__ == "__main__":
    main()
