# F04 PowerPoint→Marp変換機能 機能設計書

## 1. 概要

本機能は、PowerPoint `.pptx` ファイルを解析し、AIが内容を理解しやすく、人間も扱いやすいMarp Markdownとスライド単位 `.drawio.svg` へ変換する。

PowerPointのスライド内容は、1スライドにつき1つのdraw.io編集データ付き `.drawio.svg` として出力する。
Marp Markdownはスライド順を維持し、各 `.drawio.svg` を画像参照する表示用ラッパーとして扱う。

```text
PowerPoint .pptx
  → 内部変換データ
  → スライド別draw.io SVG生成
  → Marp Markdown + slide/*.drawio.svg
```

## 2. 変換方針

| 項目 | 方針 |
|---|---|
| 優先順位 | AIが内容を理解できること、人間が `.drawio.svg` とMarp Markdownを扱いやすいこと、Marpでプレビューできること、Git差分でレビューしやすいこと、元PowerPointに近い見た目の順に優先する |
| PowerPoint解析方式 | `.pptx` を zip / Open XML として直接解析する |
| 出力構成 | メインMarp Markdown + slide/*.drawio.svg |
| スライド | PowerPointの各スライドをMarpの1スライドへ変換する |
| スライド区切り | Marpの `---` を使用する |
| Marp本文 | 各スライドの `.drawio.svg` を参照する最小構成とする |
| タイトル、本文、表、画像、背景、図形 | スライド単位の `.drawio.svg` 内に変換・埋め込みする |
| 画像 | 個別画像ファイルとしては出力せず、`.drawio.svg` 内に埋め込む |
| 図形・図解 | スライド全体の `.drawio.svg` として `slide/` 配下へ出力し、Marp Markdownから相対参照する |
| ノート | 発表者ノートをMarp presenter notesとして出力する |
| テーマ | 初期対応ではMarp標準テーマを使用し、テーマCSSファイルは出力しない |
| 完全再現 | PowerPointの完全なレイアウト、アニメーション、画面切替は目的としない |

### 2.1 既存変換機能から流用する方針

本機能は、F02 Word→Markdown変換機能とF03 Excel→Markdown変換機能で整理したOffice解析、内部変換データ、相対パス参照の方針を流用する。
PowerPoint固有の入力構造は解析対象XMLが異なるが、Open XMLを直接解析して表示用Markdownを生成する考え方は既存機能と共通化する。

| 流用元 | 流用する考え方 | F04での適用 |
|---|---|---|
| F02 Word→Markdown | `.docx` をzip / Open XMLとして直接解析する | `.pptx` もzip / Open XMLとして直接解析し、HTMLやPDF経由の変換は基本採用しない |
| F02 Word→Markdown | 内部変換データを作成し、そこからMarkdownを生成する | PowerPoint解析結果を `presentation`、`slides`、`slide_svgs` に整理してからMarp Markdownを生成する |
| F02 Word→Markdown | 画像はbase64ではなく外部ファイルとして出力し、相対パスで参照する | PowerPointでは個別画像ではなく、画像を埋め込んだ `.drawio.svg` をスライド単位で外部参照する |
| F03 Excel→Markdown | Markdownでは座標の完全再現より意味ブロック化を優先する | PowerPointでは座標をMarkdownへ展開せず、座標表現に向くdraw.io SVGへ集約する |

### 2.2 PowerPoint固有として追加する方針

| 項目 | PowerPoint固有の方針 |
|---|---|
| スライド単位 | PowerPointの1スライドをMarpの1スライドへ対応させる |
| Marpヘッダー | Markdown先頭にMarp front matterを出力する |
| スライド区切り | スライド間に `---` を出力する |
| スライドSVG | 1スライドにつき `slide/slideNNN.drawio.svg` を1件出力する |
| 背景画像 | スライド背景として判断できる画像は `.drawio.svg` 内に背景画像として埋め込む |
| 発表者ノート | `notesSlide` はMarp presenter notesとしてHTMLコメントに出力する |
| draw.io表示 | `.drawio.svg` をMarp表示用画像として参照する |
| Marp確認 | ローカル画像参照を含むため、Marp CLI確認時は `--allow-local-files` を使用する |

### 2.3 初期実装で優先して共通化する処理

1. Office zip展開、rels解決、media抽出
2. 内部変換データの `blocks[]` 方式
3. スライド要素の内部変換データ化
4. draw.io SVG内ラベルによる最低限の装飾保持
5. draw.io SVG生成
6. `.drawio.svg` の外部ファイル出力と相対パス参照
7. 変換できない要素の警告化

F04で新規に設計する主な範囲は、PowerPointの `ppt/slides/slide*.xml`、`slideLayout`、`slideMaster`、`notesSlide` の解析と、Marpスライドとしての出力制御である。

## 3. 対象範囲

### 3.1 対象

- `.pptx` 形式のPowerPointファイル
- 複数スライド
- スライドタイトル
- 段落、改行
- 箇条書き、番号付きリストを含むテキスト領域
- テキストボックス
- 表
- 画像
- 単純な図形テキスト
- 図形、線、矢印を含む図解の `.drawio.svg` 化
- 発表者ノート

### 3.2 対象外

- `.ppt` 旧形式
- パスワード付きPowerPoint
- マクロ実行
- アニメーション、画面切替
- 音声、動画
- SmartArtの完全再現
- グループ図形の完全な編集可能オブジェクト化
- PowerPointテーマ、マスター、レイアウトの完全再現
- フォント埋め込みデータの再利用
- PowerPointと完全一致する位置、余白、重なり順の再現

対象外要素は、取得できる情報に応じて `.drawio.svg` 内への近似変換、テキスト抽出、警告の順に扱う。

## 4. 入出力

### 4.1 入力

| 入力 | 内容 |
|---|---|
| PowerPointファイル | `.pptx` |
| ファイル選択方法 | 画面上で対象PowerPointファイルを右クリックし、変換メニューから選択する |

### 4.2 出力

入力PowerPointファイルと同一フォルダに、入力ファイル名のフォルダを作成する。

```text
入力:
  /path/to/sample.pptx

出力:
  /path/to/sample/sample.md
  /path/to/sample/slide/
  /path/to/sample/slide/slide001.drawio.svg
  /path/to/sample/slide/slide002.drawio.svg
```

| 出力 | 内容 |
|---|---|
| Marp Markdown | 変換後のスライド本文。拡張子は `.md` |
| slide | Markdownから参照するスライド単位の `.drawio.svg` |

## 5. 推奨フォルダ構成

PowerPoint→Marpでは、スライド全体を編集可能な `.drawio.svg` に集約する。
出力は本文と `slide/` 配下のスライド単位SVGに絞る。

```text
sample/
  sample.md
  slide/
    slide001.drawio.svg
    slide002.drawio.svg
    slide003.drawio.svg
```

### 5.1 ファイル命名

| 種別 | 命名 |
|---|---|
| Marp Markdown | `<入力ファイル名>.md` |
| スライドSVG | `slide<3桁>.drawio.svg` |

### 5.2 Marpからの参照例

```md
---
marp: true
theme: default
paginate: true
---

![bg contain](slide/slide001.drawio.svg)

---

![bg contain](slide/slide002.drawio.svg)

---

![bg contain](slide/slide003.drawio.svg)
```

## 6. Marp内でdraw.ioを表示する方針

Marp Markdown内にdraw.io図を表示したい場合は、`.drawio` XMLを直接埋め込むのではなく、`.drawio.svg` を画像として参照する。

| 方式 | Marp表示 | draw.io編集 | 方針 |
|---|---:|---:|---|
| `.drawio` を直接参照 | 不可 | 可 | 採用しない |
| `.drawio.svg` を画像参照 | 可 | 可 | 採用 |
| `.svg` をHTMLとして直接埋め込み | 可 | 条件付き | Markdownが読みにくくなるため原則採用しない |
| draw.io HTML embed | 条件付き | 条件付き | 外部JavaScriptやネットワーク依存が出るため採用しない |

```text
slide/slide001.drawio.svg
```

Marp本文では表示用SVGを通常の画像として参照する。

```md
![w:900](slide/slide001.drawio.svg)
```

draw.ioはSVGへ図の編集データを埋め込めるため、`.drawio.svg` は「Marpで表示できる画像」でありながら「draw.ioで再編集できる図」として扱える。
SVG内の `foreignObject`、フォント、複雑なHTMLラベルは表示環境によって崩れる可能性があるため、表示崩れが起きた場合は `.drawio.svg` の生成内容または図形変換内容を修正対象とする。

## 7. 全体処理フロー

```text
ユーザーがPowerPointファイルを選択
  → .pptxをzipとして開く
  → presentation.xmlからスライド一覧を取得
  → presentation.xml.relsからslide参照を解決
  → slide*.xmlから図形、テキスト、表、画像参照を取得
  → slide*.xml.relsから画像、ノート、埋め込み参照を解決
  → slideMasters、slideLayouts、themeを必要範囲で解析
  → スライドごとに図形、テキスト、表、画像、背景をdraw.io XMLへ変換
  → スライド単位の.drawio.svgを出力
  → Marp Markdownを生成
```

## 8. Open XML解析

### 8.1 主な解析対象

| Open XMLファイル | 用途 |
|---|---|
| `ppt/presentation.xml` | スライド一覧、スライドサイズ |
| `ppt/_rels/presentation.xml.rels` | presentationからslide、theme、masterへの参照 |
| `ppt/slides/slide*.xml` | スライド内の図形、テキスト、表、画像参照 |
| `ppt/slides/_rels/slide*.xml.rels` | slideから画像、ノート、埋め込みオブジェクトへの参照 |
| `ppt/notesSlides/notesSlide*.xml` | 発表者ノート |
| `ppt/slideLayouts/slideLayout*.xml` | プレースホルダー、レイアウト情報 |
| `ppt/slideMasters/slideMaster*.xml` | 共通スタイル、背景、プレースホルダー |
| `ppt/theme/theme*.xml` | 色、フォントなどのテーマ情報 |
| `ppt/media/*` | 画像、SVGなどのバイナリ |
| `ppt/charts/chart*.xml` | グラフ情報。初期対応では対象外または警告 |
| `ppt/embeddings/*` | 埋め込みファイル。初期対応では対象外または警告 |

### 8.2 スライド本文

`ppt/slides/slide*.xml` の `p:cSld/p:spTree` 配下を順に解析し、各スライドの出力内容を決定する。
PowerPoint上の重なり順はMarkdownに向かないため、基本は上→下、左→右の位置順で出力する。

| PowerPoint要素 | 内部変換データ | Marp Markdown |
|---|---|---|
| `p:sp` | text / shape | `.drawio.svg` 内のテキストまたは図形 |
| `p:pic` | image | `.drawio.svg` 内の画像 |
| `p:graphicFrame/a:tbl` | table | `.drawio.svg` 内のHTML table相当ラベル |
| `p:graphicFrame/c:chart` | chart | `.drawio.svg` 内の画像または警告 |
| `p:grpSp` | group | `.drawio.svg` 内のグループ近似または警告 |
| `p:cxnSp` | connector | `.drawio.svg` 候補 |

## 9. 内部変換データ設計

内部変換データは、PowerPointのOpen XMLから取得した情報をMarp Markdown生成に使いやすい形へ整理したメモリ上のデータ構造である。
ファイルとしては出力しない。

### 9.1 全体項目

| 項目 | 内容 |
|---|---|
| `presentation` | PowerPoint全体の情報 |
| `slides` | スライドごとの解析結果配列 |
| `slide_svgs` | スライド単位の `.drawio.svg` 出力一覧 |
| `warnings` | 変換時の警告一覧 |

### 9.2 slides[]

| 項目 | 内容 |
|---|---|
| `index` | スライド番号 |
| `source` | zip内のslide XMLパス |
| `title` | タイトル候補 |
| `layout` | 参照するslideLayout |
| `blocks` | draw.io SVG生成対象のスライド要素配列 |
| `notes` | 発表者ノート |
| `background` | 背景情報。`.drawio.svg` 内に埋め込む |
| `warnings` | スライド固有の警告 |

### 9.3 blocks[]

| 項目 | 内容 |
|---|---|
| `type` | `title`、`paragraph`、`list`、`table`、`image`、`shape`、`connector` など |
| `shape_id` | PowerPoint上の図形ID |
| `name` | PowerPoint上の図形名 |
| `text` | 抽出したテキスト |
| `runs` | 文字列と文字装飾の配列 |
| `position` | スライド上の位置。`x`、`y`、`width`、`height` |
| `level` | 箇条書き階層 |
| `rows` | 表の場合の行データ |
| `embedded_asset` | 画像などを `.drawio.svg` 内へ埋め込むための情報 |

### 9.4 slide_svgs[]

| 項目 | 内容 |
|---|---|
| `id` | 内部スライドSVG ID |
| `slide_index` | 参照元スライド番号 |
| `output_path` | Marp表示用ファイル。`slide/slideNNN.drawio.svg` |
| `embedded_assets` | SVG内へ埋め込んだ画像、背景、SVGなど |
| `warnings` | スライドSVG生成時の警告 |

## 10. Markdown生成

### 10.1 Marpヘッダー

出力Markdownの先頭にはMarp用front matterを出力する。

```md
---
marp: true
theme: default
paginate: true
---
```

### 10.2 スライド生成

各PowerPointスライドは、Marpの1スライドとして出力する。

```md
---

![bg contain](slide/slide001.drawio.svg)
```

Marp Markdownには、スライド本文の再構築ではなく、スライド単位の `.drawio.svg` 参照を出力する。

### 10.3 テキスト・箇条書き

PowerPointの段落、改行、箇条書き、番号付きリストは、Markdown本文へ展開せず、スライド単位 `.drawio.svg` 内のテキスト要素として出力する。
箇条書き階層は、draw.ioラベル内で視認できるようにインデントまたは行頭記号として保持する。

```text
親項目
  ・子項目
```

### 10.4 表

PowerPoint表は、Markdown本文へHTML tableとして展開せず、スライド単位 `.drawio.svg` 内の表相当要素として出力する。
セル結合や装飾を保持しやすい場合は、draw.ioラベル内でHTML table相当の表現を利用する。
styleは意味を持つ範囲に限定する。

```html
<table>
  <tr>
    <th>項目</th>
    <th>内容</th>
  </tr>
  <tr>
    <td>目的</td>
    <td>変換方式の整理</td>
  </tr>
</table>
```

### 10.5 画像

画像は個別ファイルとして出力せず、スライド単位の `.drawio.svg` 内へ埋め込む。

```md
![bg contain](slide/slide004.drawio.svg)
```

スライド背景として使われている画像も、`.drawio.svg` 内の背景画像として埋め込む。

```md
![bg contain](slide/slide004.drawio.svg)
```

### 10.6 発表者ノート

発表者ノートはMarp presenter notesとして出力する。

```md
<!--
ここに発表者ノートを出力する。
-->
```

## 11. 図形・図解変換

PowerPointスライドの図形、線、矢印、画像、表、背景は、スライド単位のdraw.io XMLへ変換する。
変換結果は `.drawio.svg` として `slide/` に出力し、Marpでは `.drawio.svg` を画像として表示する。

```md
![bg contain](slide/slide001.drawio.svg)
```

スライド内で扱えない要素は、近似変換、テキスト抽出、警告の順に扱う。

## 12. Marp変換・確認コマンド

Marp CLIでHTML、PDF、PPTXへ変換する想定コマンドは以下とする。

```bash
npx @marp-team/marp-cli@latest sample.md -o sample.html --allow-local-files
npx @marp-team/marp-cli@latest sample.md -o sample.pdf --allow-local-files
npx @marp-team/marp-cli@latest sample.md -o sample.pptx --allow-local-files
```

ローカル画像、SVG、draw.io表示用ファイルを参照するため、PDF/PPTX変換では `--allow-local-files` を付与する。
入力Markdownが信頼できる場合のみ使用する。

## 13. エラー・警告

| 事象 | 方針 |
|---|---|
| `.pptx` として開けない | エラー終了 |
| パスワード付きPowerPoint | エラー終了 |
| slide参照が壊れている | 該当スライドをスキップし警告 |
| 画像参照が壊れている | `.drawio.svg` 内へ埋め込まず警告 |
| 未対応図形 | 近似変換、テキスト抽出、警告の順に扱う |
| SmartArt | 初期対応では警告またはテキスト抽出 |
| 動画、音声 | 出力せず警告 |
| フォントが再現できない | `.drawio.svg` 内のテキスト可読性を優先し警告 |

## 14. 制約

- PowerPointと完全一致するスライドレイアウトは保証しない。
- Marp Markdownは、PowerPointの自由配置、重なり順、アニメーション再現に向かない。
- ローカルリソースを使ってMarp CLIでPDF/PPTXへ変換する場合、`--allow-local-files` が必要になる。
- draw.ioはMarpへ直接埋め込まず、表示用 `.drawio.svg` として参照する。
- `.drawio.svg` は編集データを埋め込めるが、SVG表示環境によって一部ラベルやフォントが崩れる可能性がある。
- 大量の画像やSVGを含むPowerPointでは、`.drawio.svg` のファイルサイズが大きくなる。

## 15. 参考情報

- MarpはMarkdownをHTML、PDF、PowerPointへ変換できる。
- Marp CLIはローカルファイル参照をセキュリティ上デフォルトで制限しており、必要な場合は `--allow-local-files` を使用する。
- draw.ioはSVGに図の編集データを埋め込める。
- draw.ioのSVG埋め込みはWebページ上で表示できるが、`foreignObject` など表示環境依存の制約がある。
