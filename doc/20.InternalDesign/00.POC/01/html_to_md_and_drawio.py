#!/usr/bin/env python3
import base64
import os
import sys
from html.parser import HTMLParser
from pathlib import Path
from typing import List, Optional


class SpreadsheetHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.rows: List[List[str]] = []
        self.current_row: Optional[List[str]] = None
        self.current_cell: Optional[List[str]] = None
        self.image_src: Optional[str] = None
        self.image_width: Optional[int] = None
        self.image_height: Optional[int] = None
        self.in_img = False

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "table" and attrs.get("class", "") == "waffle":
            self.in_table = True
        elif self.in_table and tag == "tr":
            self.current_row = []
        elif self.in_table and tag in ("td", "th"):
            self.current_cell = []
        elif tag == "img" and self.image_src is None:
            self.image_src = attrs.get("src")
            self.image_width = int(attrs.get("width", 0)) if attrs.get("width") else None
            self.image_height = int(attrs.get("height", 0)) if attrs.get("height") else None
            self.in_img = True

    def handle_endtag(self, tag):
        if tag == "table" and self.in_table:
            self.in_table = False
        elif self.in_table and tag in ("td", "th") and self.current_cell is not None:
            text = "".join(self.current_cell).strip()
            self.current_row.append(text)
            self.current_cell = None
        elif self.in_table and tag == "tr" and self.current_row is not None:
            self.rows.append(self.current_row)
            self.current_row = None
        elif tag == "img":
            self.in_img = False

    def handle_data(self, data):
        if self.current_cell is not None:
            self.current_cell.append(data)


def make_markdown(rows: List[List[str]], image_src: Optional[str]) -> str:
    if not rows:
        return ""
    header = rows[0]
    body = rows[1:]
    sep = ["---"] * len(header)
    lines = ["| " + " | ".join(header) + " |", "| " + " | ".join(sep) + " |"]
    for row in body:
        padded = [cell if cell else "" for cell in row]
        if len(padded) < len(header):
            padded += [""] * (len(header) - len(padded))
        lines.append("| " + " | ".join(padded) + " |")
    if image_src:
        lines.append("")
        lines.append("## Embedded Diagram")
        lines.append("")
        lines.append(f"![diagram]({image_src})")
    return "\n".join(lines)


def make_drawio_xml(rows: List[List[str]], image_path: Optional[Path], image_width: int, image_height: int) -> str:
    image_data = None
    if image_path and image_path.exists():
        with open(image_path, "rb") as f:
            image_data = base64.b64encode(f.read()).decode("ascii")
    if image_data is None:
        image_style = "shape=rectangle;rounded=1;whiteSpace=wrap;html=1;"
    else:
        image_style = f"shape=image;image=data:image/png;base64,{image_data};html=1;"

    object_text = []
    if len(rows) > 1:
        headers = rows[0]
        for row in rows[1:]:
            if len(row) >= 4 and row[2] and row[3]:
                object_text.append(f"{row[2]}: {row[3]}")
    obj_text = "\\n".join(object_text) if object_text else ""
    obj_text = obj_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    cell_id = 2
    image_cell = f"<mxCell id=\"{cell_id}\" value=\"\" style=\"{image_style}\" vertex=\"1\" parent=\"1\">"
    image_geom = f"<mxGeometry x=\"20\" y=\"20\" width=\"{image_width}\" height=\"{image_height}\" as=\"geometry\"/>"
    image_cell += image_geom + "</mxCell>"

    text_cell = f"<mxCell id=\"{cell_id + 1}\" value=\"{obj_text}\" style=\"shape=rectangle;rounded=1;whiteSpace=wrap;html=1;align=left;verticalAlign=top;\" vertex=\"1\" parent=\"1\">"
    text_geom = f"<mxGeometry x=\"20\" y=\"{image_height + 40}\" width=\"380\" height=\"{max(60, 20 * len(object_text))}\" as=\"geometry\"/>"
    text_cell += text_geom + "</mxCell>"

    return f"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<mxfile host=\"app.diagrams.net\">
  <diagram name=\"Sheet-Export\">
    <mxGraphModel dx=\"1000\" dy=\"700\" grid=\"1\" gridSize=\"10\" guides=\"1\" tooltips=\"1\" connect=\"1\" arrows=\"1\" fold=\"1\" page=\"1\" pageScale=\"1\" pageWidth=\"827\" pageHeight=\"1169\" math=\"0\" shadow=\"0\">
      <root>
        <mxCell id=\"0\"/>
        <mxCell id=\"1\" parent=\"0\"/>
        {image_cell}
        {text_cell}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
"""


def main():
    if len(sys.argv) < 2:
        print("Usage: html_to_md_and_drawio.py <input-html> [output-prefix]")
        sys.exit(1)

    html_path = Path(sys.argv[1])
    if not html_path.exists():
        raise FileNotFoundError(html_path)

    output_prefix = Path(sys.argv[2]) if len(sys.argv) > 2 else html_path.with_suffix("")
    parser = SpreadsheetHTMLParser()
    parser.feed(html_path.read_text(encoding="utf-8"))

    md_text = make_markdown(parser.rows, parser.image_src)
    md_path = output_prefix.with_suffix(".md")
    md_path.write_text(md_text, encoding="utf-8")
    print(f"Wrote Markdown: {md_path}")

    image_path = None
    if parser.image_src:
        image_path = html_path.parent / parser.image_src
    image_width = parser.image_width or 700
    image_height = parser.image_height or 200
    drawio_xml = make_drawio_xml(parser.rows, image_path, image_width, image_height)
    drawio_path = output_prefix.with_suffix(".drawio")
    drawio_path.write_text(drawio_xml, encoding="utf-8")
    print(f"Wrote draw.io XML: {drawio_path}")


if __name__ == "__main__":
    main()
