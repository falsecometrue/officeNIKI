# HTML/CSS/IMG 解析レポート

## テーブル解析
- row 0:  | A | B | C | D | E
- row 1: 1 | サンプルテスト |  |  |  | 
- row 2:  |  |  |  |  | 
- row 3: 2 |  |  |  |  | 
- row 4: 3 |  |  |  |  | 
- row 5: 4 |  |  |  |  | 
- row 6: 5 |  |  |  |  | 
- row 7: 6 |  |  |  |  | 
- row 8: 7 |  |  |  |  | 
- row 9: 8 |  |  |  |  | 
- row 10: 9 |  |  |  |  | 
- row 11: 10 |  |  |  |  | 
- row 12: 11 |  |  |  |  | 
- row 13: 12 |  |  |  |  | 
- row 14: 13 |  |  |  |  | 
- row 15: 14 |  |  |  |  | 
- row 16: 15 |  |  |  |  | 
- row 17: 16 |  |  |  |  | 
- row 18: 17 |  | no | title | detail | memo
- row 19: 18 |  | 1 | オブジェクトA | 青, 中央テキスト：”オブジェクトA” | -
- row 20: 19 |  | 2 | 右矢印 | オブジェクトAからオブジェクトBへの矢印 | 細
- row 21: 20 |  | 3 | オブジェクトB | 赤, 中央テキスト：”オブジェクトB” | -
- row 22: 21 |  |  |  |  | 
- row 23: 22 |  |  |  |  | 
- row 24: 23 |  |  |  |  | 
- row 25: 24 | end |  |  |  | 

## 画像リソース
- src: `resources/drawing0.png` alt: `` width: `698` height: `215` parent: `div`

## JS 位置情報
- object_id: `embed_2068220731` sheet: `0` row: 2 col: 0 x: 98 y: 16`

## オーバーレイ/埋め込みオブジェクト
- id: `embed_2068220731` class: `['waffle-embedded-object-overlay']` style: `width: 698px; height: 215px; display: block;`

## 使用 CSS クラス
- .waffle: `font-size:13px;table-layout:fixed;border-collapse:separate;border-style:none;border-spacing:0;width:0;cursor:default`
- .grid-container: `background-color:var(--gm3-sys-color-surface-container-low,#f8fafd);overflow:hidden;position:relative;z-index:0`
- .freezebar-vertical-handle: `width:4px;background:url("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAYCAIAAABSh8vJAAAAEElEQVQYV2PYgwQYRjnEcgDquNOBEawK+wAAAABJRU5ErkJggg==") no-repeat`
- .freezebar-horizontal-handle: `height:4px;background:url("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAC4AAAAECAIAAAB+7JowAAAAEklEQVQY02PYM2gAw6hTBrdTAFI0lXC16jX6AAAAAElFTkSuQmCC") no-repeat`
- .column-headers-background: `z-index:1`
- .row-headers-background: `z-index:1`
- .row-header: `background:#f9f9f9`
- .waffle-embedded-object-overlay: `outline:0;position:absolute;z-index:10`
- .ritz: `color: inherit;`

## 解析の考察
- テーブル内容は Markdown に直接変換できます。
- 埋め込みオブジェクトは画像として出力されており、ネイティブ draw.io 図形に自動変換するには意味解釈が必要です。
- JavaScript の `posObj` からは、オブジェクトの配置ヒントを取得できます。
- CSS は外部ファイルが巨大なので、必要なクラスのみ抽出して解析しています。
