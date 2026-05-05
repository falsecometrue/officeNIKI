#!/usr/bin/env python3
import json
import re
from html.parser import HTMLParser
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple


class SpreadsheetExportParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.stack: List[Tuple[str, Dict[str, str]]] = []
        self.style_blocks: List[str] = []
        self.scripts: List[str] = []
        self.current_data: List[str] = []
        self.in_style = False
        self.in_script = False
        self.link_css: List[str] = []
        self.used_classes: Set[str] = set()
        self.table_rows: List[Dict] = []
        self.current_row: Optional[Dict] = None
        self.current_cell: Optional[Dict] = None
        self.images: List[Dict] = []
        self.overlays: List[Dict] = []

    def handle_starttag(self, tag, attrs):
        attrs = {k: v for k, v in attrs}
        self.stack.append((tag, attrs))
        if tag == 'style':
            self.in_style = True
            self.current_data = []
        elif tag == 'script':
            self.in_script = True
            self.current_data = []
        elif tag == 'link' and attrs.get('rel') == 'stylesheet' and attrs.get('href'):
            self.link_css.append(attrs['href'])
        if 'class' in attrs:
            self.used_classes.update(attrs['class'].split())
        if tag == 'table' and attrs.get('class', '') == 'waffle':
            self.current_row = None
        if tag == 'tr':
            self.current_row = {'row_index': len(self.table_rows), 'cells': []}
        if self.current_row is not None and tag in ('td', 'th'):
            self.current_cell = {
                'type': tag,
                'class': attrs.get('class', '').split(),
                'style': attrs.get('style', ''),
                'rowspan': int(attrs.get('rowspan', '1')),
                'colspan': int(attrs.get('colspan', '1')),
                'text': '',
            }
        if tag == 'img':
            self.images.append({
                'src': attrs.get('src'),
                'alt': attrs.get('alt', ''),
                'width': attrs.get('width'),
                'height': attrs.get('height'),
                'class': attrs.get('class', '').split(),
                'parent': self.stack[-2][0] if len(self.stack) >= 2 else None,
            })
        if tag == 'div':
            class_list = attrs.get('class', '').split()
            if any('overlay' in c or 'embedded-object' in c or 'waffle-embedded-object' in c for c in class_list):
                self.overlays.append({
                    'id': attrs.get('id'),
                    'class': class_list,
                    'style': attrs.get('style', ''),
                })

    def handle_endtag(self, tag):
        if tag == 'style' and self.in_style:
            self.style_blocks.append(''.join(self.current_data))
            self.in_style = False
            self.current_data = []
        elif tag == 'script' and self.in_script:
            self.scripts.append(''.join(self.current_data))
            self.in_script = False
            self.current_data = []
        elif tag == 'td' or tag == 'th':
            if self.current_cell is not None and self.current_row is not None:
                self.current_cell['text'] = self.current_cell['text'].strip()
                self.current_row['cells'].append(self.current_cell)
                self.current_cell = None
        elif tag == 'tr':
            if self.current_row is not None:
                self.table_rows.append(self.current_row)
                self.current_row = None
        self.stack.pop()

    def handle_data(self, data):
        if self.in_style or self.in_script:
            self.current_data.append(data)
        elif self.current_cell is not None:
            self.current_cell['text'] += data
        elif self.stack and self.stack[-1][0] == 'tr' and self.current_row is None:
            self.current_row = {'row_index': len(self.table_rows), 'cells': []}


def parse_css_rules(css_text: str, class_names: Set[str]) -> Dict[str, str]:
    rules: Dict[str, str] = {}
    for match in re.finditer(r'([^\{]+)\{([^\}]*)\}', css_text, re.DOTALL):
        selectors = match.group(1).strip()
        decl = match.group(2).strip()
        for selector in selectors.split(','):
            selector = selector.strip()
            if selector.startswith('.'):
                name = selector[1:].split(':')[0].split()[0]
                if name in class_names and name not in rules:
                    rules[name] = decl
    return rules


def parse_js_positions(js_text: str) -> List[Dict[str, int]]:
    positions: List[Dict[str, int]] = []
    pattern = re.compile(r"posObj\s*\(\s*(['\"])(.*?)\1\s*,\s*(['\"])(.*?)\3\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*\)")
    for m in pattern.finditer(js_text):
        positions.append({
            'sheet': int(m.group(2)) if m.group(2).isdigit() else m.group(2),
            'object_id': m.group(4),
            'row': int(m.group(5)),
            'col': int(m.group(6)),
            'x': int(m.group(7)),
            'y': int(m.group(8)),
        })
    return positions


def load_css_sources(html_path: Path, link_css: List[str], style_blocks: List[str]):
    sources: List[Tuple[str, str]] = []
    for href in link_css:
        css_file = html_path.parent / href
        if css_file.exists():
            sources.append((href, css_file.read_text(encoding='utf-8', errors='ignore')))
    for idx, block in enumerate(style_blocks):
        sources.append((f'inline-{idx}', block))
    return sources


def make_markdown(summary: Dict[str, object]) -> str:
    lines: List[str] = []
    lines.append('# HTML/CSS/IMG 解析レポート')
    lines.append('')
    if summary['table']['found']:
        lines.append('## テーブル解析')
        for row in summary['table']['rows']:
            line = ' | '.join(cell['text'] or '' for cell in row['cells'])
            lines.append(f'- row {row["row_index"]}: {line}')
        lines.append('')
    else:
        lines.append('## テーブル解析')
        lines.append('- `waffle` テーブルが見つかりませんでした。')
        lines.append('')
    lines.append('## 画像リソース')
    if summary['images']:
        for img in summary['images']:
            lines.append(f'- src: `{img["src"]}` alt: `{img["alt"]}` width: `{img["width"]}` height: `{img["height"]}` parent: `{img["parent"]}`')
    else:
        lines.append('- 画像は見つかりませんでした。')
    lines.append('')
    lines.append('## JS 位置情報')
    if summary['js_positions']:
        for pos in summary['js_positions']:
            lines.append(f'- object_id: `{pos["object_id"]}` sheet: `{pos["sheet"]}` row: {pos["row"]} col: {pos["col"]} x: {pos["x"]} y: {pos["y"]}`')
    else:
        lines.append('- `posObj(...)` 呼び出しは見つかりませんでした。')
    lines.append('')
    lines.append('## オーバーレイ/埋め込みオブジェクト')
    if summary['overlays']:
        for overlay in summary['overlays']:
            lines.append(f'- id: `{overlay["id"]}` class: `{overlay["class"]}` style: `{overlay["style"]}`')
    else:
        lines.append('- オーバーレイ要素は見つかりませんでした。')
    lines.append('')
    lines.append('## 使用 CSS クラス')
    if summary['css_rules']:
        for class_name, decl in summary['css_rules'].items():
            lines.append(f'- .{class_name}: `{decl}`')
    else:
        lines.append('- 使用中のクラスの CSS ルールは見つかりませんでした。')
    lines.append('')
    lines.append('## 解析の考察')
    lines.append('- テーブル内容は Markdown に直接変換できます。')
    lines.append('- 埋め込みオブジェクトは画像として出力されており、ネイティブ draw.io 図形に自動変換するには意味解釈が必要です。')
    lines.append('- JavaScript の `posObj` からは、オブジェクトの配置ヒントを取得できます。')
    lines.append('- CSS は外部ファイルが巨大なので、必要なクラスのみ抽出して解析しています。')
    lines.append('')
    return '\n'.join(lines)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description='Analyze spreadsheet export HTML/CSS/images structure.')
    parser.add_argument('input_html', help='HTML export file path')
    parser.add_argument('--output-prefix', default='analysis_output', help='Output prefix for generated files')
    args = parser.parse_args()

    html_path = Path(args.input_html)
    if not html_path.exists():
        raise FileNotFoundError(f'Input HTML file not found: {html_path}')

    parser = SpreadsheetExportParser()
    parser.feed(html_path.read_text(encoding='utf-8', errors='ignore'))

    css_sources = load_css_sources(html_path, parser.link_css, parser.style_blocks)
    css_rules: Dict[str, str] = {}
    for _, css_text in css_sources:
        css_rules.update(parse_css_rules(css_text, parser.used_classes))

    js_text = '\n'.join(parser.scripts)
    js_positions = parse_js_positions(js_text)

    summary = {
        'table': {'found': bool(parser.table_rows), 'rows': parser.table_rows},
        'images': parser.images,
        'overlays': parser.overlays,
        'js_positions': js_positions,
        'css_rules': css_rules,
        'css_sources': [href for href, _ in css_sources],
    }

    output_prefix = Path(args.output_prefix)
    json_path = output_prefix.with_suffix('.json')
    md_path = output_prefix.with_suffix('.md')
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    md_path.write_text(make_markdown(summary), encoding='utf-8')

    print(f'Wrote JSON analysis: {json_path}')
    print(f'Wrote Markdown analysis: {md_path}')


if __name__ == '__main__':
    main()
