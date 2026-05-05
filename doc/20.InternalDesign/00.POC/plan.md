# POC 計画: Excel/HTML → AI が読み込みやすい形式への変換

## 1. 背景
- `doc/30.test/testData/Excel_sampleData/シート1.html` は Google スプレッドシートの HTML 変換出力。
- Excel の「Web ページとして保存」や Google スプレッドシートの HTML エクスポートは、HTML/CSS/JS/画像のセットを生成する点で類似。
- 目的は、これらの出力を AI が意味を理解しやすい構造化形式に変換すること。

## 2. サンプル HTML から見えた課題
- 表は `<table class="waffle">` で出力される。
- 図形・イラストは `resources/drawing0.png` など画像リソースとして出力され、追加の JS で位置調整が行われている。
- CSS が内部 `<style>` と外部 `resources/sheet.css` に分かれる。
- そのため、単純な HTML テキスト抽出だけでは図形や位置情報を正しく把握できない。

## 3. 実現方針

### 3.1 Excel → HTML/CSS/JS/画像 の実現
1. Excel 側ではまず、既存のエクスポート機能を活用する。
   - Excel の「Web ページとして保存」/「HTML エクスポート」
   - Google スプレッドシートの HTML 変換と同等の出力を期待。
2. 自動化が必要な場合は、以下を検討する。
   - `LibreOffice` / `soffice --convert-to html` での変換
   - Python の `openpyxl` / `xlrd` などによるセル・画像抽出 + HTML 生成
   - Node.js の `sheetjs` / `exceljs` などでセルデータを取得し、HTML とリソースを生成
3. 図形や描画については、Excel の出力に含まれる以下を確認する。
   - `drawing*.png` などの画像データ
   - 画像位置を決める JavaScript
   - セルのマージやスタイル情報

### 3.2 HTML/CSS/JS/画像 → AI が読みやすい形式 の実現
#### 3.2.1 解析手順
- HTML を DOM ベースで解析
- CSS を読み込んで必要なスタイルを簡略化
- 画像リソースを収集
- JavaScript による位置調整があれば、可能な限り再現または画像ベースで扱う

#### 3.2.2 出力形式
- Markdown
  - シート名/テーブルを Markdown 表に変換
  - 画像は `![]()` 形式で埋め込み
  - 図形はテキスト化して補足説明を付与
- JSON
  - `sheets`, `rows`, `cols`, `cells`, `drawings`, `images` などの構造化データ
  - セル値、スタイル、マージ、座標、図形オブジェクトの属性を含める
- Draw.io XML
  - 図形と矢印を `mxGraphModel` 相当の XML 形式へ変換
  - 位置情報・サイズ・ラベルを持つオブジェクト列として表現

#### 3.2.3 図形・イラスト対応
- HTML 中の埋め込み画像や `<img>` を検出
- サンプルのように描画オブジェクト位置が JS で制御される場合は、以下の手順を検討
  - JS を実行してレンダリング後の位置情報を取得（Headless Chrome/Puppeteer）
  - 位置情報が取得できない場合は、画像自体を AI へ渡すための `img` パス/alt テキストを残す
- 図形を構造化するため、以下の情報を抽出
  - 図形タイプ（四角、矢印、テキストラベル）
  - 色、サイズ、座標
  - 元データの意味（例: オブジェクトA、オブジェクトB、矢印）

### 3.3 変換方針
大前提、webなどでアップロードしてGUI上で人間側で補正する工程がある。

・テーブルは直接変換  
・いったん画像はそのまま流用する。コピーなどで  
  backlog意味解釈についてはOCRなどで解釈させる？  
  半自動化で人間による調整もあるから一旦保留  

・オブジェクトヒントはより正確に見積もる方法はある？  
 　人間の補正があるからヒントのままでもよい。

### 3.4 変換テーブル
| Excel | 中間（HTML/CSS/JS） | AI リーダブル（Markdown / Draw.io） |
|---|---|---|
| 表データ | `<table class="waffle">` + CSS スタイル | Markdown 表 / JSON テーブル / Draw.io テーブル |
| セルテキスト | HTML セル内テキスト | Markdown テキスト / Draw.io ラベル |
| マージセル・スタイル | HTML 属性（rowspan/colspan） + CSS | Markdown ではセル結合を簡易表現 / Draw.io では位置・幅で再現 |
| 図形・矢印 | 画像リソース + JS 位置情報（posObj 等） | 画像埋め込み / Draw.io 図形ヒント（位置・ラベル） |
| 画像 | `<img>` / `resources/*.png` | Markdown `![]()` / Draw.io 画像要素 |
| レイアウト情報 | CSS + JS 座標 | Markdown では補足説明 / Draw.io では x,y,size |

### 4.1 変換ライブラリ候補
- Excel → HTML 生成
  - `LibreOffice` / `soffice --convert-to html`
  - `sheetjs` / `exceljs` / `xlsx`（Node.js）
  - `openpyxl` / `python-docx`（Python）
- HTML/CSS/JS 解析
  - `jsdom` / `cheerio`（Node.js）
  - `BeautifulSoup` / `lxml`（Python）
  - `Puppeteer` / `Playwright`（レンダリング後の位置取得）
- Draw.io XML 生成
  - 独自変換コードで `mxGraphModel` を構築
  - 既存の draw.io ライブラリがあれば活用

### 4.2 実装パターン
- Option A: Excel から直接構造化データへ変換
  - `.xlsx` を直接解析し、HTML を経由せず JSON/Draw.io/XML を生成
  - 図形は `xl/drawings` を解析
- Option B: Excel → HTML 出力を一次生成し、HTML を解析して構造化データへ変換
  - 本 POC のサンプルに適した方法
  - `HTML + CSS + 画像 + JS` をそのまま扱う

## 5. POC 進行計画
1. サンプル HTML の詳細解析
   - `シート1.html` の table 構造、CSS、画像埋め込み、JS 制御を調査
2. HTML 解析モジュールの作成
   - セルデータ抽出、画像収集、図形抽出
3. 出力フォーマット設計
   - Markdown、JSON、Draw.io XML のスキーマ
4. 実装と検証
   - `Excel_sampleData` を入力として、AI 読み取り形式への変換を試行
5. Excel 直接出力対応の検証
   - Excel の HTML 出力形式と同等の結果を再現できるか確認

## 6. まとめ
- `excel → html,css,js,img → AIが読み込みやすい形` は、まず既存の HTML エクスポート形式を利用し、そこから構造化変換を行う。
- `html,css,js,img → AIが読み込みやすい形` では、DOM 解析に加えて JS レンダリングや画像埋め込みを前提とした処理が必要。
- 最終的には、Markdown / JSON / Draw.io XML の3つの形式を出力できるパイプラインを確立する。
