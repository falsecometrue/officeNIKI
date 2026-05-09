#!/usr/bin/env python3
import argparse
import json
import posixpath
import re
import shutil
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr


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
SHEET_ORIGIN_X = 40
SHEET_ORIGIN_Y = 40
RESOURCE_DIR_NAME = "resources"


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
    if target.startswith("/"):
        return target.lstrip("/")
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


def display_value_from_raw(cell_type: str, raw_value: str, shared_strings: List[str]) -> str:
    if cell_type == "s" and raw_value:
        return shared_strings[int(raw_value)]
    if cell_type == "n" and raw_value:
        try:
            number = Decimal(raw_value)
        except InvalidOperation:
            return raw_value
        if number == number.to_integral_value():
            return str(number.quantize(Decimal("1")))
    return raw_value


def normalize_rgb(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    rgb = value[-6:]
    if len(rgb) != 6:
        return None
    return f"#{rgb}"


# styles.xml cellXfs are referenced from sheet cells by style_id.
def parse_styles(package: zipfile.ZipFile) -> List[Dict[str, object]]:
    if "xl/styles.xml" not in package.namelist():
        return []
    root = read_xml(package, "xl/styles.xml")
    fills = []
    for fill in root.findall("main:fills/main:fill", NS):
        fg_color = fill.find(".//main:fgColor", NS)
        fills.append({
            "color": normalize_rgb(fg_color.attrib.get("rgb")) if fg_color is not None else None,
        })
    borders = []
    for border in root.findall("main:borders/main:border", NS):
        border_color = None
        for side_name in ("left", "right", "top", "bottom"):
            side = border.find(f"main:{side_name}", NS)
            color = side.find("main:color", NS) if side is not None else None
            border_color = normalize_rgb(color.attrib.get("rgb")) if color is not None else border_color
        borders.append({"color": border_color})
    styles = []
    for index, xf in enumerate(root.findall("main:cellXfs/main:xf", NS)):
        fill_id = int(xf.attrib.get("fillId", "0"))
        border_id = int(xf.attrib.get("borderId", "0"))
        styles.append({
            "style_id": index,
            "font_id": int(xf.attrib.get("fontId", "0")),
            "fill_id": fill_id,
            "fill_color": fills[fill_id]["color"] if fill_id < len(fills) else None,
            "border_id": border_id,
            "border_color": borders[border_id]["color"] if border_id < len(borders) else None,
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
            value = display_value_from_raw(cell_type, raw_value, shared_strings)
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


def image_mime(image_path: str) -> str:
    suffix = Path(image_path).suffix.lower().lstrip(".")
    mime = "jpeg" if suffix in {"jpg", "jpeg"} else suffix or "octet-stream"
    return f"image/{mime}"


def copy_image_resource(
    package: zipfile.ZipFile,
    image_path: str,
    resource_dir: Path,
    copied_images: Dict[str, str],
) -> str:
    # POC08: base64 埋め込みではなく、Draw.io から参照する外部画像として書き出す。
    if image_path in copied_images:
        return copied_images[image_path]

    resource_dir.mkdir(parents=True, exist_ok=True)
    dest = resource_dir / Path(image_path).name
    if dest.exists():
        stem = dest.stem
        suffix = dest.suffix
        index = 2
        while (resource_dir / f"{stem}_{index}{suffix}").exists():
            index += 1
        dest = resource_dir / f"{stem}_{index}{suffix}"
    dest.write_bytes(package.read(image_path))
    rel_path = f"{RESOURCE_DIR_NAME}/{dest.name}"
    copied_images[image_path] = rel_path
    return rel_path


def excel_col_width_to_px(width: float) -> float:
    # Approximation used for layout POC. Excel's exact width depends on default font metrics.
    return round(width * 7 + 5, 2)


def excel_row_height_to_px(height_pt: float) -> float:
    return round(height_pt * PX_PER_INCH / 72, 2)


def col_width_map_from_cols(cols: List[Dict[str, object]]) -> Dict[int, float]:
    widths: Dict[int, float] = {}
    for col in cols:
        for index in range(int(col["min"]), int(col["max"]) + 1):
            widths[index] = excel_col_width_to_px(float(col.get("width") or DEFAULT_COL_WIDTH))
    return widths


def row_height_map_from_rows(rows: List[Dict[str, object]]) -> Dict[int, float]:
    heights: Dict[int, float] = {}
    for row in rows:
        height = row.get("height") or DEFAULT_ROW_HEIGHT_PT
        heights[int(row["index"])] = excel_row_height_to_px(float(height))
    return heights


def col_width_map(sheet: Dict[str, object]) -> Dict[int, float]:
    return col_width_map_from_cols(sheet.get("cols", []))


def row_height_map(sheet: Dict[str, object]) -> Dict[int, float]:
    return row_height_map_from_rows(sheet.get("rows", []))


def sheet_x_from_zero_based_col(col: int, col_widths: Dict[int, float]) -> float:
    default_col_px = excel_col_width_to_px(DEFAULT_COL_WIDTH)
    return SHEET_ORIGIN_X + sum(col_widths.get(i, default_col_px) for i in range(1, col + 1))


def sheet_y_from_zero_based_row(row: int, row_heights: Dict[int, float]) -> float:
    default_row_px = excel_row_height_to_px(DEFAULT_ROW_HEIGHT_PT)
    return SHEET_ORIGIN_Y + sum(row_heights.get(i, default_row_px) for i in range(1, row + 1))


# Anchors tell where a drawing is attached to the sheet grid.
def parse_anchor(anchor: ET.Element, col_widths: Dict[int, float], row_heights: Dict[int, float]) -> Dict[str, object]:
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
        result["x"] = round(
            sheet_x_from_zero_based_col(result["from"]["col"], col_widths)
            + emu_to_px(result["from"]["colOff_emu"]),
            2,
        )
        result["y"] = round(
            sheet_y_from_zero_based_row(result["from"]["row"], row_heights)
            + emu_to_px(result["from"]["rowOff_emu"]),
            2,
        )
    if ext is not None:
        result["ext"] = {
            "cx_emu": int(ext.attrib.get("cx", "0")),
            "cy_emu": int(ext.attrib.get("cy", "0")),
            "width": emu_to_px(ext.attrib.get("cx", "0")),
            "height": emu_to_px(ext.attrib.get("cy", "0")),
        }
    return result


def parse_group_transform(group: ET.Element) -> Dict[str, int]:
    xfrm = group.find("xdr:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return {"off_x": 0, "off_y": 0, "ch_off_x": 0, "ch_off_y": 0}
    off = xfrm.find("a:off", NS)
    ch_off = xfrm.find("a:chOff", NS)
    return {
        "off_x": int(off.attrib.get("x", "0")) if off is not None else 0,
        "off_y": int(off.attrib.get("y", "0")) if off is not None else 0,
        "ch_off_x": int(ch_off.attrib.get("x", "0")) if ch_off is not None else 0,
        "ch_off_y": int(ch_off.attrib.get("y", "0")) if ch_off is not None else 0,
    }


# Convert one DrawingML shape/connector into the intermediate drawing object.
def parse_drawing_shape(node: ET.Element, anchor_info: Dict[str, object], group_transform: Optional[Dict[str, int]] = None) -> Dict[str, object]:
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
        raw_x_emu = int(off.attrib.get("x", "0"))
        raw_y_emu = int(off.attrib.get("y", "0"))
        ch_off_x = group_transform["ch_off_x"] if group_transform else 0
        ch_off_y = group_transform["ch_off_y"] if group_transform else 0
        local_x_emu = raw_x_emu - ch_off_x
        local_y_emu = raw_y_emu - ch_off_y
        anchor_x = float(anchor_info.get("x", 0))
        anchor_y = float(anchor_info.get("y", 0))
        shape.update({
            "raw_x_emu": raw_x_emu,
            "raw_y_emu": raw_y_emu,
            "local_x_emu": local_x_emu,
            "local_y_emu": local_y_emu,
            "x_emu": local_x_emu,
            "y_emu": local_y_emu,
            "x": round(anchor_x + emu_to_px(local_x_emu), 2),
            "y": round(anchor_y + emu_to_px(local_y_emu), 2),
        })
    if ext is not None:
        shape.update({
            "width_emu": int(ext.attrib.get("cx", "0")),
            "height_emu": int(ext.attrib.get("cy", "0")),
            "width": emu_to_px(ext.attrib.get("cx", "0")),
            "height": emu_to_px(ext.attrib.get("cy", "0")),
        })
    return shape


def parse_drawing_picture(
    package: zipfile.ZipFile,
    node: ET.Element,
    anchor_info: Dict[str, object],
    drawing_rels: Dict[str, Dict[str, str]],
    drawing_path: str,
    picture_index: int,
    resource_dir: Path,
    copied_images: Dict[str, str],
) -> Dict[str, object]:
    nv = node.find(".//xdr:cNvPr", NS)
    blip = node.find(".//a:blip", NS)
    rel_id = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else ""
    target = drawing_rels.get(rel_id, {}).get("target", "")
    image_path = resolve_package_path(drawing_path, target) if target else ""
    external_path = copy_image_resource(package, image_path, resource_dir, copied_images) if image_path else ""
    ext = anchor_info.get("ext") or {}
    return {
        "id": f"pic-{picture_index}",
        "name": nv.attrib.get("name") if nv is not None else "",
        "kind": "image",
        "rel_id": rel_id,
        "path": image_path,
        "x": anchor_info.get("x", 0),
        "y": anchor_info.get("y", 0),
        "width": ext.get("width", 120),
        "height": ext.get("height", 80),
        "image_ref_type": "external_path",
        "external_path": external_path,
    }


# drawing*.xml can contain anchors, grouped shapes, connectors, and pictures.
def parse_drawings(
    package: zipfile.ZipFile,
    drawing_path: str,
    rows: List[Dict[str, object]],
    cols: List[Dict[str, object]],
    resource_dir: Path,
    copied_images: Dict[str, str],
) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    if drawing_path not in package.namelist():
        return [], []
    root = read_xml(package, drawing_path)
    rel_path = f"{posixpath.dirname(drawing_path)}/_rels/{posixpath.basename(drawing_path)}.rels"
    drawing_rels = read_relationships(package, rel_path)
    col_widths = col_width_map_from_cols(cols)
    row_heights = row_height_map_from_rows(rows)
    anchors = []
    drawings = []
    picture_index = 1
    for anchor in list(root):
        if not anchor.tag.endswith(("oneCellAnchor", "twoCellAnchor", "absoluteAnchor")):
            continue
        anchor_info = parse_anchor(anchor, col_widths, row_heights)
        anchors.append(anchor_info)
        for child in list(anchor):
            if child.tag.endswith("sp") or child.tag.endswith("cxnSp") or child.tag.endswith("pic"):
                if child.tag.endswith("pic"):
                    drawings.append(parse_drawing_picture(package, child, anchor_info, drawing_rels, drawing_path, picture_index, resource_dir, copied_images))
                    picture_index += 1
                else:
                    drawings.append(parse_drawing_shape(child, anchor_info))
            elif child.tag.endswith("grpSp"):
                group_transform = parse_group_transform(child)
                for group_child in list(child):
                    if group_child.tag.endswith("sp") or group_child.tag.endswith("cxnSp") or group_child.tag.endswith("pic"):
                        if group_child.tag.endswith("pic"):
                            drawings.append(parse_drawing_picture(package, group_child, anchor_info, drawing_rels, drawing_path, picture_index, resource_dir, copied_images))
                            picture_index += 1
                        else:
                            drawings.append(parse_drawing_shape(group_child, anchor_info, group_transform))
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


# Keep media as external files so the Draw.io file does not contain large base64 blobs.
def read_images(package: zipfile.ZipFile, resource_dir: Path, copied_images: Dict[str, str]) -> List[Dict[str, object]]:
    images = []
    for name in package.namelist():
        if not name.startswith("xl/media/"):
            continue
        external_path = copy_image_resource(package, name, resource_dir, copied_images)
        images.append({
            "path": name,
            "mime": image_mime(name),
            "image_ref_type": "external_path",
            "external_path": external_path,
        })
    return images


# Main xlsx -> intermediate JSON pipeline.
def build_intermediate(xlsx_path: Path, output_dir: Path) -> Dict[str, object]:
    resource_dir = output_dir / RESOURCE_DIR_NAME
    if resource_dir.exists():
        shutil.rmtree(resource_dir)
    copied_images: Dict[str, str] = {}
    with zipfile.ZipFile(xlsx_path) as package:
        shared_strings = read_shared_strings(package)
        styles = parse_styles(package)
        sheet_defs = parse_workbook(package)
        images = read_images(package, resource_dir, copied_images)
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
                    drawing_items, anchor_items = parse_drawings(package, drawing_path, rows, cols, resource_dir, copied_images)
                    drawings.extend(drawing_items)
                    anchors.extend(anchor_items)
            sheets.append({
                "name": sheet_def["name"],
                "sheet_id": sheet_def["sheet_id"],
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
                "image_output": "external_path",
                "resources_dir": RESOURCE_DIR_NAME,
                "sheets_count": len(sheets),
            },
            "sheets": sheets,
        }


def drawio_shape_style(shape: Dict[str, object]) -> str:
    fill = shape.get("fill") or "#FFFFFF"
    stroke = shape.get("stroke") or "#000000"
    rounded = "1;arcSize=12" if shape.get("preset") == "flowChartAlternateProcess" else "0"
    return f"rounded={rounded};whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};"


def cell_geometry(cell: Dict[str, object], col_widths: Dict[int, float], row_heights: Dict[int, float]) -> Dict[str, float]:
    col = int(cell["col"])
    row = int(cell["row"])
    default_col_px = excel_col_width_to_px(DEFAULT_COL_WIDTH)
    default_row_px = excel_row_height_to_px(DEFAULT_ROW_HEIGHT_PT)
    x = SHEET_ORIGIN_X + sum(col_widths.get(i, default_col_px) for i in range(1, col))
    y = SHEET_ORIGIN_Y + sum(row_heights.get(i, default_row_px) for i in range(1, row))
    return {
        "x": round(x, 2),
        "y": round(y, 2),
        "width": col_widths.get(col, default_col_px),
        "height": row_heights.get(row, default_row_px),
    }


def drawio_cell_style(cell: Dict[str, object], styles: Dict[int, Dict[str, object]]) -> str:
    style = styles.get(int(cell.get("style_id", 0)), {})
    if is_table_cell(cell, styles):
        fill = style.get("fill_color") or "#FFFFFF"
        stroke = style.get("border_color") or "#000000"
        return (
            "rounded=0;whiteSpace=wrap;html=1;"
            f"fillColor={fill};strokeColor={stroke};"
            "align=center;verticalAlign=middle;fontSize=11;"
        )
    return (
        "text;html=1;"
        "fillColor=none;strokeColor=none;"
        "align=center;verticalAlign=middle;fontSize=11;"
    )


def is_table_cell(cell: Dict[str, object], styles: Dict[int, Dict[str, object]]) -> bool:
    style = styles.get(int(cell.get("style_id", 0)), {})
    return bool(style.get("apply_border") or style.get("fill_color"))


def find_table_components(cells: List[Dict[str, object]], styles: Dict[int, Dict[str, object]]) -> List[List[Dict[str, object]]]:
    table_cells = {
        (int(cell["row"]), int(cell["col"])): cell
        for cell in cells
        if is_table_cell(cell, styles)
    }
    visited = set()
    components: List[List[Dict[str, object]]] = []
    for key in table_cells:
        if key in visited:
            continue
        stack = [key]
        visited.add(key)
        component = []
        while stack:
            row, col = stack.pop()
            component.append(table_cells[(row, col)])
            for neighbor in ((row - 1, col), (row + 1, col), (row, col - 1), (row, col + 1)):
                if neighbor in table_cells and neighbor not in visited:
                    visited.add(neighbor)
                    stack.append(neighbor)
        components.append(component)
    return components


def table_geometry(component: List[Dict[str, object]], col_widths: Dict[int, float], row_heights: Dict[int, float]) -> Dict[str, float]:
    rows = [int(cell["row"]) for cell in component]
    cols = [int(cell["col"]) for cell in component]
    min_row, max_row = min(rows), max(rows)
    min_col, max_col = min(cols), max(cols)
    top_left = cell_geometry({"row": min_row, "col": min_col}, col_widths, row_heights)
    width = sum(col_widths.get(col, excel_col_width_to_px(DEFAULT_COL_WIDTH)) for col in range(min_col, max_col + 1))
    height = sum(row_heights.get(row, excel_row_height_to_px(DEFAULT_ROW_HEIGHT_PT)) for row in range(min_row, max_row + 1))
    return {
        "x": top_left["x"],
        "y": top_left["y"],
        "width": round(width, 2),
        "height": round(height, 2),
        "min_row": min_row,
        "max_row": max_row,
        "min_col": min_col,
        "max_col": max_col,
    }


def html_td_style(cell: Optional[Dict[str, object]], styles: Dict[int, Dict[str, object]], width: float, height: float) -> str:
    style = styles.get(int(cell.get("style_id", 0)), {}) if cell else {}
    fill = style.get("fill_color") or "#FFFFFF"
    stroke = style.get("border_color") or "#000000"
    return (
        f"border:1px solid {stroke};"
        f"background:{fill};"
        "text-align:center;"
        "vertical-align:middle;"
        f"width:{width}px;"
        f"height:{height}px;"
        "font-size:11px;"
        "padding:0;"
        "box-sizing:border-box;"
    )


def build_html_table(component: List[Dict[str, object]], styles: Dict[int, Dict[str, object]], col_widths: Dict[int, float], row_heights: Dict[int, float]) -> str:
    geometry = table_geometry(component, col_widths, row_heights)
    by_position = {
        (int(cell["row"]), int(cell["col"])): cell
        for cell in component
    }
    lines = [
        '<table style="border-collapse:collapse;table-layout:fixed;width:100%;height:100%;font-family:Arial,sans-serif;">'
    ]
    for row in range(int(geometry["min_row"]), int(geometry["max_row"]) + 1):
        row_height = row_heights.get(row, excel_row_height_to_px(DEFAULT_ROW_HEIGHT_PT))
        lines.append("<tr>")
        for col in range(int(geometry["min_col"]), int(geometry["max_col"]) + 1):
            col_width = col_widths.get(col, excel_col_width_to_px(DEFAULT_COL_WIDTH))
            cell = by_position.get((row, col))
            value = escape(str(cell.get("value", ""))) if cell else ""
            style = html_td_style(cell, styles, col_width, row_height)
            lines.append(f'<td style="{style}">{value}</td>')
        lines.append("</tr>")
    lines.append("</table>")
    return "".join(lines)


def append_html_table_vertices(
    xml_cells: List[str],
    sheet: Dict[str, object],
    styles: Dict[int, Dict[str, object]],
    col_widths: Dict[int, float],
    row_heights: Dict[int, float],
) -> set:
    table_cell_refs = set()
    for index, component in enumerate(find_table_components(sheet.get("cells", []), styles), start=1):
        if not component:
            continue
        geometry = table_geometry(component, col_widths, row_heights)
        html = build_html_table(component, styles, col_widths, row_heights)
        table_id = f'table-{int(geometry["min_row"])}-{int(geometry["min_col"])}-{int(geometry["max_row"])}-{int(geometry["max_col"])}'
        value = quoteattr(html)
        style = quoteattr("html=1;whiteSpace=wrap;overflow=fill;rounded=0;fillColor=none;strokeColor=none;")
        xml_cells.append(f'        <mxCell id="{table_id}" value={value} style={style} vertex="1" parent="1">')
        xml_cells.append(f'          <mxGeometry x="{geometry["x"]}" y="{geometry["y"]}" width="{geometry["width"]}" height="{geometry["height"]}" as="geometry"/>')
        xml_cells.append("        </mxCell>")
        table_cell_refs.update(str(cell["ref"]) for cell in component)
    return table_cell_refs


def append_cell_vertices(xml_cells: List[str], sheet: Dict[str, object]) -> None:
    styles = {int(style["style_id"]): style for style in sheet.get("styles", [])}
    col_widths = col_width_map(sheet)
    row_heights = row_height_map(sheet)
    table_cell_refs = append_html_table_vertices(xml_cells, sheet, styles, col_widths, row_heights)
    for cell in sheet.get("cells", []):
        if str(cell["ref"]) in table_cell_refs:
            continue
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


def append_image_vertices(xml_cells: List[str], sheet: Dict[str, object]) -> None:
    for image in sheet.get("drawings", []):
        if image.get("kind") != "image":
            continue
        cell_id = f"image-{image['id']}"
        style = quoteattr(f"shape=image;html=1;imageAspect=0;aspect=fixed;image={image.get('external_path', '')};")
        xml_cells.append(f'        <mxCell id="{cell_id}" value="" style={style} vertex="1" parent="1">')
        xml_cells.append(f'          <mxGeometry x="{image.get("x", 0)}" y="{image.get("y", 0)}" width="{image.get("width", 120)}" height="{image.get("height", 80)}" as="geometry"/>')
        xml_cells.append("        </mxCell>")


def drawio_diagram_xml(sheet: Dict[str, object], index: int) -> List[str]:
    sheet_name = str(sheet.get("name") or f"Sheet{index}")
    diagram_id = quoteattr(f"POC08-{index}")
    diagram_name = quoteattr(sheet_name)
    cells = [
        f"  <diagram id={diagram_id} name={diagram_name}>",
        '    <mxGraphModel dx="1000" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169">',
        '      <root>',
        '        <mxCell id="0"/>',
        '        <mxCell id="1" parent="0"/>',
    ]
    append_cell_vertices(cells, sheet)
    append_image_vertices(cells, sheet)
    append_drawing_vertices(cells, sheet)
    append_drawing_edges(cells, sheet)
    cells.extend([
        "      </root>",
        "    </mxGraphModel>",
        "  </diagram>",
    ])
    return cells


# Intermediate JSON -> Draw.io XML. Each Excel sheet becomes one Draw.io page.
def make_drawio(intermediate: Dict[str, object]) -> str:
    cells = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mxfile host="app.diagrams.net">',
    ]
    for index, sheet in enumerate(intermediate.get("sheets", []), start=1):
        cells.extend(drawio_diagram_xml(sheet, index))
    cells.extend([
        "</mxfile>",
    ])
    return "\n".join(cells)


# CLI entrypoint for repeatable POC runs.
def main() -> None:
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parents[4]
    default_input = repo_root / "doc" / "30.test" / "00.pocTestData" / "テストデータ.xlsx"
    if not default_input.exists():
        default_input = repo_root / "doc" / "30.test" / "testData" / "テストデータ.xlsx"
    parser = argparse.ArgumentParser(description="POC08: xlsx zip XML to Draw.io with external image paths")
    parser.add_argument("--input", type=Path, default=default_input)
    parser.add_argument("--json", type=Path, default=script_dir / "poc08_intermediate.json")
    parser.add_argument("--drawio", type=Path, default=script_dir / "poc08_output.drawio")
    args = parser.parse_args()

    args.json.parent.mkdir(parents=True, exist_ok=True)
    args.drawio.parent.mkdir(parents=True, exist_ok=True)
    intermediate = build_intermediate(args.input, args.drawio.parent)
    args.json.write_text(json.dumps(intermediate, ensure_ascii=False, indent=2), encoding="utf-8")
    args.drawio.write_text(make_drawio(intermediate), encoding="utf-8")
    print(f"Wrote {args.json}")
    print(f"Wrote {args.drawio}")


if __name__ == "__main__":
    main()
