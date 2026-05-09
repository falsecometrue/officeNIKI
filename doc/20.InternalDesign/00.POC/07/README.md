# 07: Excel表のDraw.io表現方式比較

## 1. 目的

POC06では、Excelの表を「セルごとの `mxCell`」としてDraw.ioへ出力した。

この方式は座標・罫線・セル単位編集には強い一方、Draw.io上では1つの表オブジェクトではなく、セルオブジェクトの集合になる。

本メモでは、次の2方式を比較する。

- HTML tableラベル方式
- Draw.io専用table形状方式

比較対象は、Excelから抽出した中間JSONの以下の情報をDraw.ioへどう反映するかである。

- `cells`
- `rows`
- `cols`
- `merged_cells`
- `styles`
- `x`, `y`, `width`, `height`

## 2. 比較サマリ

| 観点 | HTML tableラベル方式 | Draw.io専用table形状方式 |
|---|---|---|
| 表としてのまとまり | 強い。1つの `mxCell` にできる | 強い可能性がある |
| セル結合 | `colspan` / `rowspan` で表現可能 | 方式次第。要検証 |
| セル単位の編集 | 弱い。HTMLラベル内部の編集になる | 方式次第 |
| セル単位の移動 | 不可に近い | 方式次第 |
| 見た目再現 | 中から高。CSS次第 | 中から高の可能性 |
| Excel座標再現 | 表全体の座標は簡単。セルごとの座標はHTML内で表現 | 方式次第 |
| 罫線/背景色 | CSSで表現可能 | style定義で表現できる可能性 |
| 行高/列幅 | `width`, `height`, `style` で近似可能 | 方式次第 |
| Draw.io互換性 | HTMLレンダリング仕様に依存 | Draw.io内部仕様に依存 |
| 実装難易度 | 低から中 | 中から高 |
| POC向き | 高い | 調査向き |

現時点の推奨は、**短期はHTML tableラベル方式を追加検証し、既存のセルごとの `mxCell` 方式と切り替え可能にする**ことである。

Draw.io専用table形状方式は、Draw.ioの保存XMLを実際に作成して解析し、構造が安定しているかを確認してから採用判断する。

## 3. HTML tableラベル方式

### 3.1 概要

Excelの表範囲をHTMLの `<table>` として組み立て、そのHTML文字列をDraw.ioの1つの `mxCell.value` に埋め込む方式。

Draw.io XMLのイメージ:

```xml
<mxCell id="table-B17-E20"
        value="&lt;table border=&quot;1&quot; cellspacing=&quot;0&quot; cellpadding=&quot;4&quot;&gt;...&lt;/table&gt;"
        style="html=1;whiteSpace=wrap;overflow=fill;"
        vertex="1"
        parent="1">
  <mxGeometry x="133.41" y="376" width="710.48" height="84" as="geometry"/>
</mxCell>
```

### 3.2 セル結合

Excelの `merged_cells` はHTMLの `colspan` / `rowspan` に変換できる。

Excel:

```text
A1:C1
```

HTML:

```html
<td colspan="3">結合セル</td>
```

縦方向の結合:

```html
<td rowspan="2">縦結合</td>
```

縦横結合:

```html
<td colspan="3" rowspan="2">縦横結合</td>
```

### 3.3 スタイル

セル単位の背景色、罫線、文字寄せは `td` のstyleへ変換する。

例:

```html
<td style="border:1px solid #000000;background:#FFFFFF;text-align:center;vertical-align:middle;width:168px;height:21px;">
  title
</td>
```

Excelからの対応:

| Excel情報 | HTML table |
|---|---|
| `cells[].value` | `td` のテキスト |
| `cols[].width` | `td` / `col` の `width` |
| `rows[].height` | `tr` / `td` の `height` |
| `merged_cells` | `colspan`, `rowspan` |
| `styles[].fill_color` | `background` |
| `styles[].border_color` | `border` |
| `alignment` | `text-align`, `vertical-align` |

### 3.4 メリット

- Draw.io上で表が1つのまとまりになる。
- セル結合を表構造として表現しやすい。
- XMLの要素数が少なくなる。
- 表全体を移動・拡大縮小しやすい。
- Markdown/HTMLへの再利用もしやすい。

### 3.5 デメリット

- Draw.io上でセル単位にクリックして動かす用途には向かない。
- HTMLラベル内部の編集体験は、Excelの表編集とは異なる。
- Draw.ioのHTMLレンダリングでCSSがどこまで反映されるか検証が必要。
- 図形・画像とセル単位で重ね合わせる用途には弱い。

### 3.6 向いている用途

- Excel表を「表として読む」ことを優先する。
- Draw.io上では表全体を動かせればよい。
- セル結合や罫線を見た目として再現したい。
- 表をAIリーダブルなHTML構造として残したい。

## 4. Draw.io専用table形状方式

### 4.1 概要

Draw.ioが持つテーブル系・コンテナ系の形状を使って、Excel表をDraw.ioネイティブに近い表として表現する方式。

候補:

- Draw.ioのtable系shape
- swimlane/table風shape
- コンテナ + row/cell子要素

ただし、Draw.ioのtable形状は通常の `mxCell` と比べて内部XML仕様に依存しやすい。採用するには、Draw.io上で実際に表を作り、保存XMLを解析する必要がある。

### 4.2 想定XML

正確なXMLは実測が必要だが、方向性としては以下のような構造になる可能性がある。

```xml
<mxCell id="table-1" value="" style="shape=table;html=1;" vertex="1" parent="1">
  <mxGeometry x="133.41" y="376" width="710.48" height="84" as="geometry"/>
</mxCell>
```

または、表コンテナの配下に行・セルを子 `mxCell` として持つ形になる可能性がある。

```xml
<mxCell id="table-1" value="" style="swimlane;html=1;" vertex="1" parent="1">
  <mxGeometry x="133.41" y="376" width="710.48" height="84" as="geometry"/>
</mxCell>
<mxCell id="table-1-row-1" parent="table-1" ... />
<mxCell id="table-1-cell-1-1" parent="table-1-row-1" ... />
```

### 4.3 セル結合

セル結合をDraw.io専用table形状で表現できるかは未確定。

確認すべき点:

- 横結合が可能か
- 縦結合が可能か
- 結合セルのXML表現が安定しているか
- Draw.io UIで編集したときに構造が維持されるか
- diagrams.netのバージョン差異に強いか

### 4.4 メリット

- Draw.io上で表オブジェクトらしく扱える可能性がある。
- UI上の表編集に近い操作ができる可能性がある。
- HTMLラベルよりDraw.ioネイティブな見た目になる可能性がある。

### 4.5 デメリット

- Draw.io内部仕様への依存が強い。
- XML生成ルールの調査コストが高い。
- Excelのセル結合、罫線、列幅、行高をどこまで表現できるか未確定。
- 複雑な表で崩れた場合、原因調査が難しくなる可能性がある。

### 4.6 向いている用途

- Draw.io上で表として再編集したい。
- Draw.ioのUI操作と相性のよい表現を優先したい。
- HTMLラベルではなく、Draw.ioネイティブな部品として扱いたい。

## 5. 既存方式との位置づけ

POC06の既存方式は「セルごとの `mxCell` 方式」である。

| 方式 | 位置づけ |
|---|---|
| セルごとの `mxCell` | Excelの見た目・座標・セル単位補正を優先 |
| HTML tableラベル | 表としてのまとまり・セル結合・HTML構造を優先 |
| Draw.io専用table形状 | Draw.io上での表編集体験を優先できる可能性 |

POC06の方式を完全に置き換えるのではなく、出力モードとして分けるのがよい。

```text
--table-mode cells
--table-mode html
--table-mode drawio-table
```

## 6. 実装案

### 6.1 中間JSONの拡張

表を扱いやすくするため、`cells` の集合から表範囲を明示した `tables` を追加する。

```json
{
  "tables": [
    {
      "id": "table-B17-E20",
      "range": "B17:E20",
      "x": 133.41,
      "y": 376,
      "width": 710.48,
      "height": 84,
      "rows": [17, 18, 19, 20],
      "cols": [2, 3, 4, 5],
      "cells": ["B17", "C17", "D17", "E17"],
      "merged_cells": []
    }
  ]
}
```

初期実装では、連続した罫線付きセル範囲を表として推定する。

### 6.2 HTML table生成

1. `tables[].range` のセル行列を作る。
2. `merged_cells` を見て `colspan` / `rowspan` を決める。
3. 結合に吸収されるセルは出力しない。
4. `styles` から `td style` を生成する。
5. HTMLをXMLエスケープして `mxCell.value` に入れる。

### 6.3 Draw.io専用table形状検証

1. Draw.io UIで以下の表を手作成する。
   - 通常表
   - 横結合あり
   - 縦結合あり
   - 背景色/罫線あり
2. `.drawio` XMLを保存する。
3. `mxCell` 構造を解析する。
4. 生成可能性を判断する。

## 7. 推奨方針

短期:

- POC06に `--table-mode html` を追加する。
- 既存の `cells` モードは残す。
- HTML tableラベル方式で `B17:E20` のような表範囲を1つの `mxCell` として出力する。
- `merged_cells` がある場合は `colspan` / `rowspan` を使う。

中期:

- Draw.io UIで専用table形状のXMLを採取する。
- セル結合、スタイル、行列追加の挙動を検証する。
- XML構造が安定していれば `--table-mode drawio-table` を追加する。

現時点では、**HTML tableラベル方式が次に検証すべき第一候補**である。

理由:

- セル結合を素直に表現できる。
- 実装コストが低い。
- 中間JSONの `tables` と相性がよい。
- 失敗しても既存のセルごとの `mxCell` 方式へ戻しやすい。
