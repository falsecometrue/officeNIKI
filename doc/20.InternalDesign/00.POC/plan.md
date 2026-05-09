# POC 計画: Word/Excel → Markdown / Draw.io への変換

## 1. 背景
- 変換方針を `Word → Markdown`、`Excel → Draw.io` に変更する。
- Word は文書構造を Markdown として扱いやすくする。
- Excel は表、画像、図形、矢印、座標を Draw.io 上で確認・補正しやすくする。
- 変換時は、入力ファイルを直接最終形式へ変換せず、まず中間言語として JSON に解析結果を集約する。

```text
Word .docx  → intermediate.json → Markdown
Excel .xlsx → intermediate.json → Draw.io XML
```

## 2. 目的
- Word 文書を Markdown に変換し、AI が読みやすく、人間も差分確認しやすい形式にする。
- Excel ブックを Draw.io に変換し、表・図形・画像をGUI上で補正できる形式にする。
- 中間 JSON を残すことで、解析結果の検証、再生成、デバッグをしやすくする。

## 3. 実現方針

### 3.1 Word → Markdown
- `.docx` を解析し、見出し、段落、箇条書き、表、画像、注記などを抽出する。
- 抽出結果はまず中間 JSON に保存する。
- 中間 JSON から Markdown を生成する。
- 画像は Markdown から参照できるリソースとして出力する。

### 3.2 Excel → Draw.io
- `.xlsx` または HTML エクスポートを解析し、シート、行、列、セル、スタイル、画像、図形、矢印、座標を抽出する。
- 抽出結果はまず中間 JSON に保存する。
- 中間 JSON から Draw.io XML を生成する。
- 画像は Draw.io XML 内に base64 data URI として埋め込み、1ファイルで再現できる形式を基本とする。
- Excel 図形や矢印は、可能な範囲で Draw.io の `mxCell` オブジェクトへ変換する。
- 変換不能な複雑図形は画像としてフォールバックする。

### 3.3 中間 JSON
中間 JSON は最終成果物ではなく、変換処理の中心となる中間言語として扱う。

Word 用 JSON:
- `document`
- `sections`
- `paragraphs`
- `headings`
- `lists`
- `tables`
- `images`
- `styles`

Excel 用 JSON:
- `workbook`
- `sheets`
- `rows`
- `cols`
- `cells`
- `styles`
- `merged_cells`
- `drawings`
- `images`
- `anchors`

中間 JSON に含めるべき情報:
- テキスト値
- スタイル
- 行高・列幅
- セル結合
- 画像パスまたはbase64参照
- 図形種別
- 座標とサイズ
- リレーション情報

## 4. Excel POC の現状メモ
- `doc/30.test/testData/Excel_sampleData/シート1.html` は Google スプレッドシートの HTML 変換出力。
- Excel の「Web ページとして保存」や Google スプレッドシートの HTML エクスポートは、HTML/CSS/JS/画像のセットを生成する点で類似。
- 表は `<table class="waffle">` で出力される。
- 図形・イラストは `resources/drawing0.png` など画像リソースとして出力され、追加の JS で位置調整が行われている。
- CSS が内部 `<style>` と外部 `resources/sheet.css` に分かれる。
- そのため、単純な HTML テキスト抽出だけでは図形や位置情報を正しく把握できない。

## 5. Draw.io 変換方針
- テーブルは Draw.io の HTML table またはセルグリッドとして変換する。
- 画像は Draw.io XML の `shape=image` として出力する。
- 画像データは `data:image/png,...` の base64 data URI として埋め込む。
- 図形は Draw.io の vertex に変換する。
- 矢印は Draw.io の edge に変換する。
- 座標とサイズは Draw.io の `mxGeometry` に変換する。
- webなどへアップロードしてGUI上で人間が補正する工程を前提に、編集しやすい構造を優先する。

## 6. 変換テーブル
| 入力 | 中間 JSON | 最終出力 |
|---|---|---|
| Word 見出し/段落 | `headings`, `paragraphs` | Markdown 見出し/本文 |
| Word 箇条書き | `lists` | Markdown list |
| Word 表 | `tables` | Markdown 表 |
| Word 画像 | `images` | Markdown 画像参照 |
| Excel 表データ | `sheets`, `rows`, `cols`, `cells` | Draw.io テーブル / セルグリッド |
| Excel セルテキスト | `cells[].value` | Draw.io ラベル |
| Excel マージセル・スタイル | `merged_cells`, `styles` | Draw.io の位置・幅・スタイル |
| Excel 図形 | `drawings` | Draw.io vertex |
| Excel 矢印 | `drawings`, `anchors` | Draw.io edge |
| Excel 画像 | `images` | Draw.io 画像要素(base64) |
| レイアウト情報 | `x`, `y`, `width`, `height` | Draw.io `mxGeometry` |

## 7. 変換ライブラリ候補

### 7.1 Word → Markdown
- `python-docx`
- `pandoc`
- `mammoth`

### 7.2 Excel → 中間 JSON / Draw.io
- `openpyxl` / `xlrd`（Python）
- `sheetjs` / `exceljs` / `xlsx`（Node.js）
- `LibreOffice` / `soffice --convert-to html`

### 7.3 HTML/CSS/JS 解析
- `BeautifulSoup` / `lxml`（Python）
- `jsdom` / `cheerio`（Node.js）
- `Puppeteer` / `Playwright`（レンダリング後の位置取得）

### 7.4 Draw.io XML 生成
- 独自変換コードで `mxGraphModel` を構築する。
- 既存の draw.io ライブラリがあれば活用する。

## 8. 実装パターン

### Option A: Word から直接中間 JSON へ変換
- `.docx` を解析し、見出し、段落、表、画像を中間 JSON にする。
- 中間 JSON から Markdown を生成する。

### Option B: Excel から直接中間 JSON へ変換
- `.xlsx` を直接解析し、HTML を経由せず中間 JSON を生成する。
- 図形は `xl/drawings` を解析する。
- 中間 JSON から Draw.io XML を生成する。

### Option C: Excel → HTML 出力を解析して中間 JSON へ変換
- 本 POC のサンプルに適した方法。
- `HTML + CSS + 画像 + JS` を扱い、中間 JSON を生成する。
- 中間 JSON から Draw.io XML を生成する。

## 9. POC 進行計画
1. Word → Markdown の変換方針を整理する。
   - `.docx` の見出し、段落、表、画像を中間 JSON にする。
   - 中間 JSON から Markdown を生成する。
2. Excel → Draw.io の変換方針を整理する。
   - `.xlsx` またはHTMLエクスポートから、表・画像・図形・座標を中間 JSON にする。
   - 中間 JSON から Draw.io XML を生成する。
3. ExcelサンプルHTMLの詳細解析を継続する。
   - `シート1.html` の table 構造、CSS、画像埋め込み、JS 制御を調査する。
4. 中間 JSON スキーマを設計する。
   - Word用JSON、Excel用JSON、Draw.io生成用JSONの項目を定義する。
5. 実装と検証を行う。
   - Wordサンプルを入力として Markdown への変換を試行する。
   - `Excel_sampleData` を入力として Draw.io への変換を試行する。

## 10. まとめ
- 変換方針は `Word → Markdown`、`Excel → Draw.io` とする。
- どちらの変換でも、まず中間言語として JSON を生成し、最終成果物は JSON から生成する。
- 中間 JSON は、解析結果の検証・再生成・デバッグに使う。
- Excel → Draw.io では、画像は base64 で埋め込み、図形・矢印は可能な範囲で Draw.io オブジェクトへ変換する。
- 最終的には、Word用Markdown生成パイプラインと、Excel用Draw.io生成パイプラインを確立する。
