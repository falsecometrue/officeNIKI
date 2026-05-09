# POC01: Word → Markdown 変換調査

## 対象

- 入力: `doc/30.test/00.pocTestData/ペット プロフィール - スペアミント (1).docx`
- 出力先: `doc/20.InternalDesign/00.POC/word→md/01`
- 方針: `.docx` を zip として開き、OpenXML を中間 JSON に解析してから Markdown を生成する。

## 成果物

| ファイル | 内容 |
|---|---|
| `poc01_docx_to_md.py` | docx → 中間 JSON → Markdown の POC 実装 |
| `poc01_intermediate.json` | 抽出した段落、見出し、画像、図形候補 |
| `poc01_output.md` | Markdown 変換結果 |
| `resources/` | Markdown から参照する画像 |

## サンプル docx の内部構造

`word/document.xml` を確認した結果、本文は主に段落で構成されている。

| 種別 | 状況 |
|---|---|
| 見出し | `Title`, `Subtitle`, `Heading1` として取得可能 |
| 段落 | `w:p` から本文テキストを取得可能 |
| 表 | 今回のサンプルには `w:tbl` はなし。実装上は Markdown 表へ変換する処理を用意 |
| 画像 | `word/media/image*.png/jpg` と relationship から取得可能 |
| 図形 | `w:drawing` 内に DrawingML として存在。Markdown への直接変換は困難 |

画像リソースは以下が含まれていた。

| 画像 | 用途の推定 |
|---|---|
| `image2.jpg` | ペット写真 |
| `image1.png` | セクション区切りの小さい装飾画像 |
| `image3.png` | 図形フォールバックらしき 1x1 画像。Markdown には出力しない |

## 図の扱い

Word 図形は Draw.io のように `mxCell` へ直接変換する前提にはしない。
Markdown の出力先では座標、角丸、線、重なり、グループ化を保持しづらいため、以下の 2 方式を比較対象にする。

| 方式 | 内容 | メリット | リスク |
|---|---|---|---|
| (a) 画像として貼り付ける | Word 上の図をレンダリング画像として Markdown に貼る | 見た目の再現性が高い | 編集性が低い。docx 内の生画像だけでは図形全体画像が取れない場合がある |
| (b) Mermaid で表現する | 図形内テキストや関係を解析して Mermaid にする | Markdown 上で編集しやすい | 位置、色、線種、矢印方向の推定が必要。複雑図形は再現困難 |

今回のサンプルでは、図形部分から `オブジェクトA` と `オブジェクトB` のテキストは取得できる。
ただし矢印や配置は docx XML だけでは確定しにくいため、Mermaid は「候補」として中間 JSON に残す。

```mermaid
flowchart LR
  A[オブジェクトA] --> B[オブジェクトB]
```

## 変換方針

1. `.docx` を zip として読み込む。
2. `word/_rels/document.xml.rels` から画像 relationship を解決する。
3. `word/document.xml` から段落、見出し、表、画像、図形候補を抽出する。
4. 抽出結果を `poc01_intermediate.json` に保存する。
5. JSON から `poc01_output.md` を生成する。

## 判断

- Word → Markdown は、文章構造、画像、表を中心に変換するのが現実的。
- 図形は Markdown の主目的から外れるため、まずは画像フォールバックを基本にする。
- 単純なフロー図だけ Mermaid 化候補を出すと、Markdown の編集性を残せる。
- 図形の見た目再現を重視する場合は、LibreOffice などでページまたは図形範囲をレンダリングして画像化する追加工程が必要。
