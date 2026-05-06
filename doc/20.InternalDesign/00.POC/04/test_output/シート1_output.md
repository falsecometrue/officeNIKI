# POC04: HTML → AI リーダブル変換結果

## 1. 抽出テーブル

|  | A | B | C | D | E |
| --- | --- | --- | --- | --- | --- |
| 1 | サンプルテスト |  |  |  |  |
|  |  |  |  |  |  |
| 2 |  |  |  |  |  |
| 3 |  |  |  |  |  |
| 4 |  |  |  |  |  |
| 5 |  |  |  |  |  |
| 6 |  |  |  |  |  |
| 7 |  |  |  |  |  |
| 8 |  |  |  |  |  |
| 9 |  |  |  |  |  |
| 10 |  |  |  |  |  |
| 11 |  |  |  |  |  |
| 12 |  |  |  |  |  |
| 13 |  |  |  |  |  |
| 14 |  |  |  |  |  |
| 15 |  |  |  |  |  |
| 16 |  |  |  |  |  |
| 17 |  | no | title | detail | memo |
| 18 |  | 1 | オブジェクトA | 青, 中央テキスト：”オブジェクトA” | - |
| 19 |  | 2 | 右矢印 | オブジェクトAからオブジェクトBへの矢印 | 細 |
| 20 |  | 3 | オブジェクトB | 赤, 中央テキスト：”オブジェクトB” | - |
| 21 |  |  |  |  |  |
| 22 |  |  |  |  |  |
| 23 |  |  |  |  |  |
| 24 | end |  |  |  |  |

## 2. 画像リソース

- src: `resources/drawing0.png` copied_src: `copied_images/resources/drawing0.png` alt: `` width: `698` height: `215` parent: `div`
  ![resources/drawing0.png](copied_images/resources/drawing0.png)

## 3. オーバーレイ/埋め込みオブジェクト

- id: `embed_2068220731` class: `waffle-embedded-object-overlay` style: `width: 698px; height: 215px; display: block;` width: `698` height: `215` x: `98` y: `16` row: `2` col: `0` src: `resources/drawing0.png` copied_src: `copied_images/resources/drawing0.png`

## 4. JS 位置情報 (posObj)

- object_id: `embed_2068220731` sheet: `0` row: 2 col: 0 x: 98 y: 16`

## 5. Draw.io ヒント生成

- 埋め込みオブジェクトの座標と画像を Draw.io に反映しています。
- 可能であれば `drawing0.png` をそのまま表示する運用を想定しています。

- 表データから図形オブジェクト候補を抽出しました。
  - オブジェクトA (想定サイズ 220x90)
  - オブジェクトB (想定サイズ 220x90)

Draw.io XML は `poc04_drawio_hints.dio` / `poc04_drawio_hints.svg.dio` に出力しています。