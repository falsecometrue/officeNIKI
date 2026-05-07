#!/usr/bin/env python3
import base64
import json
import re
import shutil
from html.parser import HTMLParser
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape


class POC05Parser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.stack: List[Tuple[str, Dict[str, str]]] = []
        self.in_script = False
        self.current_data: List[str] = []
        self.table_rows: List[List[str]] = []
        self.current_row: Optional[List[str]] = None
        self.current_cell: Optional[str] = None
        self.images: List[Dict[str, Optional[str]]] = []
        self.overlays: List[Dict[str, Optional[str]]] = []
        self.scripts: List[str] = []
        self.pos_objs: List[Dict[str, object]] = []

    def handle_starttag(self, tag, attrs):
        attrs = {k: v for k, v in attrs}
        self.stack.append((tag, attrs))
        if tag == 'script':
            self.in_script = True
            self.current_data = []
        if tag == 'tr':
            self.current_row = []
        if tag in ('td', 'th') and self.current_row is not None:
            self.current_cell = ''
        if tag == 'img':
            overlay_id = None
            for parent_tag, parent_attrs in reversed(self.stack[:-1]):
                if parent_tag == 'div':
                    parent_class = parent_attrs.get('class', '')
                    if 'embedded-object' in parent_class or 'overlay' in parent_class or 'waffle-embedded-object' in parent_class:
                        overlay_id = parent_attrs.get('id')
                        break
            self.images.append({
                'src': attrs.get('src'),
                'alt': attrs.get('alt', ''),
                'width': attrs.get('width'),
                'height': attrs.get('height'),
                'class': attrs.get('class', ''),
                'parent': self.stack[-2][0] if len(self.stack) >= 2 else None,
                'overlay_id': overlay_id,
            })
        if tag == 'div':
            class_list = attrs.get('class', '')
            if 'embedded-object' in class_list or 'overlay' in class_list or 'waffle-embedded-object' in class_list:
                self.overlays.append({
                    'id': attrs.get('id'),
                    'class': class_list,
                    'style': attrs.get('style', ''),
                })

    def handle_endtag(self, tag):
        if tag == 'script' and self.in_script:
            self.scripts.append(''.join(self.current_data))
            self.in_script = False
            self.current_data = []
        if tag in ('td', 'th') and self.current_cell is not None and self.current_row is not None:
            self.current_row.append(self.current_cell.strip())
            self.current_cell = None
        if tag == 'tr' and self.current_row is not None:
            self.table_rows.append(self.current_row)
            self.current_row = None
        self.stack.pop()

    def handle_data(self, data):
        if self.in_script:
            self.current_data.append(data)
        elif self.current_cell is not None:
            self.current_cell += data

    def extract_pos_objs(self) -> None:
        pattern = re.compile(r"posObj\s*\(\s*(['\"])(.*?)\1\s*,\s*(['\"])(.*?)\3\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*\)")
        for script in self.scripts:
            for match in pattern.finditer(script):
                self.pos_objs.append({
                    'sheet': int(match.group(2)) if match.group(2).isdigit() else match.group(2),
                    'object_id': match.group(4),
                    'row': int(match.group(5)),
                    'col': int(match.group(6)),
                    'x': int(match.group(7)),
                    'y': int(match.group(8)),
                })


def table_to_markdown(rows: List[List[str]]) -> str:
    if not rows:
        return 'テーブルが見つかりませんでした。'
    width = max(len(r) for r in rows)
    normalized = [r + [''] * (width - len(r)) for r in rows]
    lines = ['| ' + ' | '.join(normalized[0]) + ' |',
             '| ' + ' | '.join(['---'] * width) + ' |']
    for row in normalized[1:]:
        lines.append('| ' + ' | '.join(row) + ' |')
    return '\n'.join(lines)


def table_to_html_label(rows: List[List[str]]) -> str:
    if not rows:
        return ''
    lines = ['<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">']
    for row in rows:
        cells = ''.join(f'<td>{escape(cell or "")}</td>' for cell in row)
        lines.append(f'<tr>{cells}</tr>')
    lines.append('</table>')
    return ''.join(lines)


def parse_style_dimensions(style: str) -> Tuple[Optional[int], Optional[int]]:
    width = None
    height = None
    if style:
        width_match = re.search(r'width\s*:\s*(\d+)px', style)
        height_match = re.search(r'height\s*:\s*(\d+)px', style)
        if width_match:
            width = int(width_match.group(1))
        if height_match:
            height = int(height_match.group(1))
    return width, height


def parse_style_position(style: str) -> Tuple[Optional[int], Optional[int]]:
    x = None
    y = None
    if style:
        left_match = re.search(r'left\s*:\s*(-?\d+)px', style)
        top_match = re.search(r'top\s*:\s*(-?\d+)px', style)
        if left_match:
            x = int(left_match.group(1))
        if top_match:
            y = int(top_match.group(1))
    return x, y


def get_png_dimensions(image_path: Path) -> Tuple[Optional[int], Optional[int]]:
    if not image_path.exists():
        return None, None
    try:
        with image_path.open('rb') as f:
            header = f.read(24)
            if len(header) != 24 or header[:8] != b'\x89PNG\r\n\x1a\n':
                return None, None
            width = int.from_bytes(header[16:20], 'big')
            height = int.from_bytes(header[20:24], 'big')
            return width, height
    except Exception:
        return None, None


def build_overlay_objects(parser_obj: POC05Parser, html_dir: Path) -> List[Dict[str, object]]:
    objects: List[Dict[str, object]] = []
    for overlay in parser_obj.overlays:
        width, height = parse_style_dimensions(overlay.get('style', ''))
        left, top = parse_style_position(overlay.get('style', ''))
        image = next((img for img in parser_obj.images if img.get('overlay_id') == overlay.get('id')), None)
        src = image.get('src') if image else None
        image_path = html_dir / src if src else None
        image_dimensions = (None, None)
        if image_path and image_path.exists():
            image_dimensions = get_png_dimensions(image_path)
        if width is None and image and image.get('width') and str(image.get('width')).isdigit():
            width = int(image.get('width'))
        if height is None and image and image.get('height') and str(image.get('height')).isdigit():
            height = int(image.get('height'))
        if width is None and image_dimensions[0] is not None:
            width = image_dimensions[0]
        if height is None and image_dimensions[1] is not None:
            height = image_dimensions[1]
        pos = next((p for p in parser_obj.pos_objs if p['object_id'] == overlay.get('id')), None)
        x = pos['x'] if pos else left
        y = pos['y'] if pos else top
        objects.append({
            'id': overlay.get('id'),
            'class': overlay.get('class'),
            'style': overlay.get('style'),
            'width': width,
            'height': height,
            'x': x,
            'y': y,
            'row': pos['row'] if pos else None,
            'col': pos['col'] if pos else None,
            'src': src,
            'style_left': left,
            'style_top': top,
            'image_actual_width': image_dimensions[0],
            'image_actual_height': image_dimensions[1],
        })
    return objects


def encode_image_data_for_drawio(image_path: Path) -> Optional[str]:
    if not image_path.exists():
        return None
    data = image_path.read_bytes()
    b64 = base64.b64encode(data).decode('ascii')
    # Draw.io style parser uses ';' as separators, so omit ';base64' here.
    return f'data:image/png,{b64}'


def make_drawio_xml(rows: List[List[str]], overlay_objects: List[Dict[str, object]], image_values: Dict[str, str]) -> str:
    table_geometry = {
        'x': 100,
        'y': 40,
        'width': 520,
        'height': 280,
    }
    xml = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mxfile host="app.diagrams.net">',
        '  <diagram id="POC05" name="Page-1">',
        '    <mxGraphModel dx="1000" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169">',
        '      <root>',
        '        <mxCell id="0"/>',
        '        <mxCell id="1" parent="0"/>',
    ]
    cell_id = 2

    if rows:
        label_html = table_to_html_label(rows)
        style = 'rounded=0;whiteSpace=wrap;html=1;fillColor=none;strokeColor=none;'
        escaped_html = escape(label_html, {'"': '&quot;'})
        xml.append(f'        <mxCell id="{cell_id}" value="{escaped_html}" style="{style}" vertex="1" parent="1">')
        xml.append(
            f'          <mxGeometry x="{table_geometry["x"]}" y="{table_geometry["y"]}" '
            f'width="{table_geometry["width"]}" height="{table_geometry["height"]}" as="geometry"/>'
        )
        xml.append('        </mxCell>')
        cell_id += 1

    for overlay in overlay_objects:
        src = overlay.get('src')
        image_value = image_values.get(src) if src else None
        if not image_value:
            continue
        x = table_geometry['x']
        y = table_geometry['y']
        width = table_geometry['width']
        height = table_geometry['height']
        image_style = f'shape=image;image={image_value};html=1;verticalLabelPosition=bottom;verticalAlign=top;'
        xml.append(f'        <mxCell id="{cell_id}" value="" style="{image_style}" vertex="1" parent="1">')
        xml.append(f'          <mxGeometry x="{x}" y="{y}" width="{width}" height="{height}" as="geometry"/>')
        xml.append('        </mxCell>')
        cell_id += 1

    xml.extend([
        '      </root>',
        '    </mxGraphModel>',
        '  </diagram>',
        '</mxfile>',
    ])
    return '\n'.join(xml)


def infer_objects_from_table(rows: List[List[str]]) -> List[Dict[str, str]]:
    objects = []
    for row in rows:
        if len(row) >= 4 and row[2].strip() in ('1', '3') and row[3].strip() in ('オブジェクトA', 'オブジェクトB'):
            objects.append({
                'label': row[3].strip(),
                'width': '220',
                'height': '90',
            })
    return objects


def copy_overlay_images(
    overlay_objects: List[Dict[str, object]], html_dir: Path, output_dir: Path
) -> Dict[str, str]:
    copied: Dict[str, str] = {}
    copied_dir = output_dir / 'copied_images'
    copied_dir.mkdir(parents=True, exist_ok=True)

    for overlay in overlay_objects:
        src = overlay.get('src')
        if not src or not isinstance(src, str) or src in copied:
            continue

        src_path = Path(src)
        if src_path.is_absolute() or '..' in src_path.parts:
            continue

        original = html_dir / src_path
        if not original.exists():
            continue

        destination = copied_dir / src_path
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(original, destination)
        copied[src] = destination.relative_to(output_dir).as_posix()

    return copied


def build_debug_log(parser_obj: POC05Parser, overlay_objects: List[Dict[str, object]]) -> str:
    lines: List[str] = []
    lines.append('=== POC05 Debug Log ===')
    lines.append('')
    lines.append('[Images]')
    if parser_obj.images:
        for idx, img in enumerate(parser_obj.images, start=1):
            lines.append(
                f'{idx}. src={img.get("src")} alt={img.get("alt")} width={img.get("width")} height={img.get("height")} '
                f'class={img.get("class")} parent={img.get("parent")} overlay_id={img.get("overlay_id")}'
            )
    else:
        lines.append('(none)')
    lines.append('')

    lines.append('[Overlay divs]')
    if parser_obj.overlays:
        for idx, overlay in enumerate(parser_obj.overlays, start=1):
            lines.append(
                f'{idx}. id={overlay.get("id")} class={overlay.get("class")} style={overlay.get("style")}'
            )
    else:
        lines.append('(none)')
    lines.append('')

    lines.append('[JS posObj]')
    if parser_obj.pos_objs:
        for idx, item in enumerate(parser_obj.pos_objs, start=1):
            lines.append(
                f'{idx}. object_id={item.get("object_id")} sheet={item.get("sheet")} row={item.get("row")} col={item.get("col")} '
                f'x={item.get("x")} y={item.get("y")}'
            )
    else:
        lines.append('(none)')
    lines.append('')

    lines.append('[Resolved overlay objects]')
    if overlay_objects:
        for idx, obj in enumerate(overlay_objects, start=1):
            lines.append(
                f'{idx}. id={obj.get("id")} src={obj.get("src")} width={obj.get("width")} height={obj.get("height")} '
                f'x={obj.get("x")} y={obj.get("y")} row={obj.get("row")} col={obj.get("col")} '
                f'copied_src={obj.get("copied_src")} '
                f'style_left={obj.get("style_left")} style_top={obj.get("style_top")} '
                f'image_actual_width={obj.get("image_actual_width")} image_actual_height={obj.get("image_actual_height")}'
            )
    else:
        lines.append('(none)')

    lines.append('')
    return '\n'.join(lines)


def process_file(html_path: Path, output_dir: Path) -> None:
    if not html_path.exists():
        raise FileNotFoundError(f'Input HTML file not found: {html_path}')

    parser_obj = POC05Parser()
    parser_obj.feed(html_path.read_text(encoding='utf-8', errors='ignore'))
    parser_obj.extract_pos_objs()

    overlay_objects = build_overlay_objects(parser_obj, html_path.parent)
    print('DEBUG overlay object resolution:')
    for obj in overlay_objects:
        print(
            f'  id={obj["id"]} src={obj["src"]} width={obj["width"]} height={obj["height"]} x={obj["x"]} y={obj["y"]} '
            f'row={obj["row"]} col={obj["col"]} style_left={obj["style_left"]} style_top={obj["style_top"]} '
            f'image_actual_width={obj["image_actual_width"]} image_actual_height={obj["image_actual_height"]}'
        )

    copied_images = copy_overlay_images(overlay_objects, html_path.parent, output_dir)
    for overlay in overlay_objects:
        src = overlay.get('src')
        overlay['copied_src'] = copied_images.get(src)

    debug_log = build_debug_log(parser_obj, overlay_objects)
    debug_log_path = html_path.parent / f'{html_path.stem}_debugLog.txt'
    debug_log_path.write_text(debug_log, encoding='utf-8')

    image_data_uris: Dict[str, str] = {}
    for overlay in overlay_objects:
        src = overlay.get('src')
        if src and src not in image_data_uris:
            copied_src = copied_images.get(src)
            image_path = output_dir / copied_src if copied_src else html_path.parent / src
            image_data_uri = encode_image_data_for_drawio(image_path)
            if image_data_uri:
                image_data_uris[src] = image_data_uri

    summary = {
        'table_rows': parser_obj.table_rows,
        'images': parser_obj.images,
        'overlays': parser_obj.overlays,
        'overlay_objects': overlay_objects,
        'pos_objs': parser_obj.pos_objs,
    }

    markdown = [
        '# POC05: HTML → AI リーダブル変換結果',
        '',
        '## 1. 抽出テーブル',
        '',
        table_to_markdown(parser_obj.table_rows),
        '',
        '## 2. 画像リソース',
        '',
    ]
    if parser_obj.images:
        for img in parser_obj.images:
            copied_src = copied_images.get(img['src'])
            markdown.append(
                f'- src: `{img["src"]}` copied_src: `{copied_src}` alt: `{img["alt"]}` '
                f'width: `{img["width"]}` height: `{img["height"]}` parent: `{img["parent"]}`'
            )
            if copied_src:
                label = img['alt'] or img['src']
                markdown.append(f'  ![{label}]({copied_src})')
    else:
        markdown.append('- 画像は見つかりませんでした。')
    markdown.append('')
    markdown.append('## 3. オーバーレイ/埋め込みオブジェクト')
    markdown.append('')
    if overlay_objects:
        for overlay in overlay_objects:
            markdown.append(
                f'- id: `{overlay["id"]}` class: `{overlay["class"]}` style: `{overlay["style"]}` '
                f'width: `{overlay["width"]}` height: `{overlay["height"]}` x: `{overlay["x"]}` y: `{overlay["y"]}` '
                f'row: `{overlay["row"]}` col: `{overlay["col"]}` src: `{overlay["src"]}` copied_src: `{overlay["copied_src"]}`'
            )
    else:
        markdown.append('- 埋め込みオブジェクトは見つかりませんでした。')
    markdown.append('')
    markdown.append('## 4. JS 位置情報 (posObj)')
    markdown.append('')
    if parser_obj.pos_objs:
        for item in parser_obj.pos_objs:
            markdown.append(f'- object_id: `{item["object_id"]}` sheet: `{item["sheet"]}` row: {item["row"]} col: {item["col"]} x: {item["x"]} y: {item["y"]}`')
    else:
        markdown.append('- posObj 情報は見つかりませんでした。')
    markdown.append('')

    objects = infer_objects_from_table(parser_obj.table_rows)
    markdown.append('## 5. Draw.io ヒント生成')
    markdown.append('')
    if overlay_objects:
        markdown.append('- 埋め込みオブジェクトの座標と画像を Draw.io に反映しています。')
        markdown.append('- 可能であれば `drawing0.png` をそのまま表示する運用を想定しています。')
        markdown.append('')
    if objects:
        markdown.append('- 表データから図形オブジェクト候補を抽出しました。')
        for obj in objects:
            markdown.append(f'  - {obj["label"]} (想定サイズ {obj["width"]}x{obj["height"]})')
        markdown.append('')
        markdown.append('Draw.io XML は `*_drawio_hints.dio` / `*_drawio_hints.xml` / `*_drawio_hints.svg.xml` に出力しています。')
    else:
        markdown.append('- 図形候補は表データから抽出できませんでした。')
        markdown.append('- 埋め込みオブジェクト画像はそのまま保持し、GUI で補正する運用を想定します。')

    stem = html_path.stem
    md_path = output_dir / f'{stem}_output.md'
    json_path = output_dir / f'{stem}_summary.json'
    drawio_base = output_dir / f'{stem}_drawio_hints'
    drawio_paths = [
        drawio_base.with_suffix('.dio'),
        drawio_base.with_suffix('.svg.xml'),
    ]
    drawio_linked_path = output_dir / f'{stem}_drawio_hints.xml'

    md_path.write_text('\n'.join(markdown), encoding='utf-8')
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    drawio_xml = make_drawio_xml(parser_obj.table_rows, overlay_objects, image_data_uris)
    drawio_linked_xml = make_drawio_xml(parser_obj.table_rows, overlay_objects, image_data_uris)
    for path in drawio_paths:
        path.write_text(drawio_xml, encoding='utf-8')
    drawio_linked_path.write_text(drawio_linked_xml, encoding='utf-8')

    print(f'Wrote POC Markdown: {md_path}')
    print(f'Wrote POC JSON: {json_path}')
    print(f'Wrote debug log: {debug_log_path}')
    if copied_images:
        print('Copied image assets:')
        for original_src, copied_src in copied_images.items():
            print(f'  - {original_src} -> {copied_src}')
    print('Wrote Draw.io hints:')
    for path in drawio_paths:
        print(f'  - {path}')
    print(f'  - {drawio_linked_path}')


def main() -> None:
    input_dir = Path('/Users/sakiyama/Desktop/excel→md/doc/30.test/testData/Excel_sampleData')
    output_dir = Path('/Users/sakiyama/Desktop/excel→md/doc/20.InternalDesign/00.POC/05/test_output')

    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f'Input directory not found: {input_dir}')

    output_dir.mkdir(parents=True, exist_ok=True)

    html_files = list(input_dir.glob('*.html'))
    if not html_files:
        raise ValueError(f'No HTML files found in {input_dir}')

    for html_file in html_files:
        process_file(html_file, output_dir)


if __name__ == '__main__':
    main()
