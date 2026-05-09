# 06: Excel xlsx zip解析からDraw.io変換方針

## 1. 調査対象

- 入力ファイル: `doc/30.test/testData/テストデータ.xlsx`
- 目的: Excelファイルをzipとして展開し、XMLを中間JSONへ構造解析したうえで、Draw.io XMLへ変換できるかを確認する。
- 結論: このサンプルでは、HTML出力経由よりも `.xlsx` のzip内部XMLを直接読む方式が適している。

## 2. xlsxをzipとして扱った結果

`.xlsx` は実体としてはzip形式のOpen XMLパッケージであるため、拡張子を `.zip` に変更しなくても `unzip` や `zipfile` で直接読める。

確認したzip内ファイル:

```text
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/worksheets/sheet1.xml
xl/worksheets/_rels/sheet1.xml.rels
xl/drawings/drawing1.xml
xl/sharedStrings.xml
xl/styles.xml
xl/theme/theme1.xml
[Content_Types].xml
_rels/.rels
```

今回のサンプルには `xl/media/*` が存在しないため、画像ファイルは含まれていない。図形と矢印は `xl/drawings/drawing1.xml` にDrawingMLとして格納されている。

## 3. XMLごとの役割

| ファイル | 主な役割 | 中間JSONでの格納先 |
|---|---|---|
| `xl/workbook.xml` | ブックとシート名の定義 | `workbook`, `sheets[].name` |
| `xl/_rels/workbook.xml.rels` | workbookからworksheet等への参照 | `workbook.relationships` |
| `xl/worksheets/sheet1.xml` | セル、行、列、描画参照 | `rows`, `cols`, `cells`, `anchors` |
| `xl/worksheets/_rels/sheet1.xml.rels` | sheetからdrawingへの参照 | `sheets[].relationships` |
| `xl/sharedStrings.xml` | 共有文字列テーブル | `cells[].value` の解決元 |
| `xl/styles.xml` | セル罫線、フォント、塗り、配置 | `styles` |
| `xl/drawings/drawing1.xml` | 図形、矢印、アンカー、座標 | `drawings`, `anchors` |

## 4. サンプルから抽出できた構造

### 4.1 シート情報

- シート数: 1
- シート名: `シート1`
- 表データは主に `A1`, `B17:E20`, `A24` に存在する。
- `C:D:E` 列には明示的な列幅が設定されている。
- マージセルは今回のサンプルでは未検出。

### 4.2 セルテキスト

`sheet1.xml` のセル値は、文字列の場合 `sharedStrings.xml` のインデックスとして保存されている。

例:

| セル | 値 |
|---|---|
| `A1` | `サンプルテスト` |
| `B17:E17` | `no`, `title`, `detail`, `memo` |
| `C18` | `オブジェクトA` |
| `D18` | `青, 中央テキスト：”オブジェクトA”` |
| `C19` | `右矢印` |
| `D19` | `オブジェクトAからオブジェクトBへの矢印` |
| `C20` | `オブジェクトB` |
| `D20` | `赤, 中央テキスト：”オブジェクトB”` |

### 4.3 図形と矢印

`drawing1.xml` には、1つの `oneCellAnchor` の中にグループ図形 `grpSp` があり、その配下に以下が存在する。

| DrawingML要素 | id | 種別 | テキスト | 変換先 |
|---|---:|---|---|---|
| `xdr:sp` | 3 | `flowChartAlternateProcess` | `オブジェクトA` | Draw.io vertex |
| `xdr:cxnSp` | 4 | `straightConnector1` | なし | Draw.io edge |
| `xdr:sp` | 5 | `flowChartAlternateProcess` | `オブジェクトB` | Draw.io vertex |

図形の塗り色:

- `オブジェクトA`: `#CFE2F3`
- `オブジェクトB`: `#EA9999`
- 矢印/枠線: `#000000`

矢印は `tailEnd type="triangle"` を持つため、Draw.ioでは `endArrow=classic` または向き補正後の `startArrow=classic` として表現する必要がある。DrawingML上では `a:endCxn id="5"` により、接続先が `Shape 5` であることを確認できる。一方、開始側接続は明示されていないため、座標から `Shape 3` を推定する。

## 5. 中間JSONスキーマ案

今回のサンプルでは、以下の単位で中間JSONに保持するとDraw.io生成へつなげやすい。

```json
{
  "workbook": {
    "source": "doc/30.test/testData/テストデータ.xlsx",
    "format": "xlsx-zip"
  },
  "sheets": [
    {
      "name": "シート1",
      "rows": [],
      "cols": [],
      "cells": [],
      "styles": [],
      "merged_cells": [],
      "drawings": [],
      "images": [],
      "anchors": []
    }
  ]
}
```

補足:

- `cells[].value` は `sharedStrings.xml` を解決した後の値を入れる。
- `cells[].style_id` は `styles.xml` の `cellXfs` 参照として残す。
- `drawings[].x_emu`, `drawings[].y_emu`, `drawings[].width_emu`, `drawings[].height_emu` は元データとして保持する。
- `drawings[].x`, `drawings[].y`, `drawings[].width`, `drawings[].height` はDraw.io生成用のpx系座標として保持する。
- グループ図形配下の座標は、親グループのオフセットを加味する必要がある。

サンプルの具体的な中間JSON例は `sample_intermediate.json` に保存する。

## 6. plan.mdの変換テーブルに基づく変換方針

| 入力 | 中間JSON | Draw.io変換方針 |
|---|---|---|
| Excel 表データ | `sheets`, `rows`, `cols`, `cells` | セルごとに `mxCell vertex=1` を生成し、表グリッドとして配置する。POC初期は表全体をHTMLラベル化してもよい。 |
| Excel セルテキスト | `cells[].value` | Draw.ioセルの `value` に設定する。改行や特殊文字はXMLエスケープする。 |
| Excel マージセル・スタイル | `merged_cells`, `styles` | マージ範囲からセル幅/高さを合算し、罫線や塗りを `style` に変換する。今回のサンプルではマージセルなし。 |
| Excel 図形 | `drawings` | `xdr:sp` を `mxCell vertex=1` に変換する。`flowChartAlternateProcess` はDraw.ioの近似図形へマッピングする。 |
| Excel 矢印 | `drawings`, `anchors` | `xdr:cxnSp` を `mxCell edge=1` に変換する。接続先idと座標からsource/targetを決定する。 |
| Excel 画像 | `images` | `xl/media/*` をbase64化し、`shape=image;image=data:image/...` のvertexに変換する。今回のサンプルでは画像なし。 |
| レイアウト情報 | `x`, `y`, `width`, `height` | `mxGeometry x/y/width/height` へ設定する。元のEMU値もデバッグ用に保持する。 |

## 7. Draw.io XML生成時の対応案

### 7.1 図形

`xdr:sp` は以下のような `mxCell` にする。

```xml
<mxCell id="shape-3" value="オブジェクトA" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#CFE2F3;strokeColor=#000000;" vertex="1" parent="1">
  <mxGeometry x="41.21" y="19.28" width="263.28" height="213.13" as="geometry"/>
</mxCell>
```

`flowChartAlternateProcess` を厳密に再現する場合は、Draw.io側の対応shapeを調査し、近似できない場合は `rounded=0` などの基本図形へフォールバックする。

### 7.2 矢印

`xdr:cxnSp` は以下のような `mxCell edge=1` にする。

```xml
<mxCell id="edge-4" value="" style="edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;strokeColor=#000000;endArrow=classic;" edge="1" parent="1" source="shape-3" target="shape-5">
  <mxGeometry relative="1" as="geometry"/>
</mxCell>
```

今回のサンプルでは `endCxn id="5"` だけが明示されているため、sourceは線分座標の始点に近い図形として推定する。

### 7.3 表

表は2段階で進める。

1. POC段階: `cells` からDraw.io上にセルグリッドを生成する。
2. 改善段階: 罫線、列幅、行高、セル結合、塗り色を `styles` と `merged_cells` から再現する。

今回のサンプルでは表データが説明表として機能しているため、図形変換の検証には `drawing1.xml` を主軸にする。

## 8. 実装上の注意点

- `.xlsx` を物理的に `.zip` へリネームする必要はない。Pythonなら `zipfile.ZipFile(xlsx_path)` で直接読める。
- `sharedStrings.xml` がないExcelもあり得るため、文字列解決は存在チェックが必要。
- すべての図形が `oneCellAnchor` とは限らない。`twoCellAnchor`, `absoluteAnchor` も処理対象にする。
- グループ図形 `grpSp` は、子図形の座標系と親オフセットを合成する必要がある。
- Excel DrawingMLの矢印方向とDraw.ioの `startArrow` / `endArrow` は向きがずれる可能性があるため、実ファイルをdraw.ioで開いて確認する。
- 複雑な図形、SmartArt、グラフは編集可能な図形化が難しい場合がある。その場合は画像化フォールバックを許容する。

## 9. POC06 Python

`poc06_xlsx_zip_to_drawio.py` は、`テストデータ.xlsx` をzipとして直接読み、以下を生成する。

- `poc06_intermediate.json`: 実解析結果の中間JSON
- `poc06_output.drawio`: 中間JSONから生成したDraw.io XML

実行方法:

```bash
python3 doc/20.InternalDesign/00.POC/06/poc06_xlsx_zip_to_drawio.py
```

任意の入力/出力を指定する場合:

```bash
python3 doc/20.InternalDesign/00.POC/06/poc06_xlsx_zip_to_drawio.py \
  --input doc/30.test/testData/テストデータ.xlsx \
  --json doc/20.InternalDesign/00.POC/06/poc06_intermediate.json \
  --drawio doc/20.InternalDesign/00.POC/06/poc06_output.drawio
```

## 10. POC結論

`doc/30.test/testData/テストデータ.xlsx` では、zip内部XMLから以下を直接取得できた。

- セル表データ: `sheet1.xml` + `sharedStrings.xml`
- セルスタイル: `styles.xml`
- 図形と矢印: `drawing1.xml`
- シートから図形への参照: `sheet1.xml.rels`

したがって、ExcelからDraw.ioへの変換は、HTML出力を解析するよりも `.xlsx` をzipとして読み、XMLを中間JSONへ変換する方式を優先する。HTML出力は、レンダリング後の見た目確認や画像フォールバック用途に限定するのがよい。
