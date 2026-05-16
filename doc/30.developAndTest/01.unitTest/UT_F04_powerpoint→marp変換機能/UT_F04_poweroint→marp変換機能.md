# UT_F04 PowerPoint→Marp変換機能 テスト仕様書

## 1. 目的

PowerPoint→Marp変換機能について、PowerPoint文書の主要要素、Marp Markdown出力、スライド単位 `.drawio.svg` 出力、フォルダ一括変換、処理時間の観点で確認する。
変換結果は、AIがスライド内容を理解しやすい構造を優先し、次に人間が `.drawio.svg` とMarp Markdownを扱いやすく、Marpでプレビューできることを確認する。

## 2. テスト方針

- 変換対象は `.pptx` とする。
- 出力は入力ファイルと同一フォルダに作成した入力PowerPointファイル名フォルダ内とする。
- 中間JSONファイルは出力しない。
- PowerPointの各スライドは、`slide/slideNNN.drawio.svg` として1件ずつ出力する。
- 画像、SVG画像、背景画像は個別ファイルとして出力せず、スライド単位 `.drawio.svg` 内へ埋め込む。
- Marp Markdownからはスライド単位 `.drawio.svg` を相対パスで参照する。
- PowerPointの各スライドはMarpの1スライドとして出力する。
- スライド区切りはMarpの `---` を使用する。
- 図形、線、矢印を含む図解はスライド単位 `.drawio.svg` 内に変換する。
- draw.ioは `.drawio` を直接Marpへ埋め込まず、`.drawio.svg` をMarp表示用画像として参照する。
- PowerPoint表示との完全一致は判定対象外とし、本仕様書に記載したMarp Markdown構造、スライドSVGファイル数、リンク数、処理時間を合格基準とする。

## 3. テスト観点

| 観点 | 確認内容 | 期待結果 |
|---|---|---|
| Marpヘッダー | front matter | Markdown先頭に `marp: true`、`theme:`、`paginate:` が出力される |
| スライド区切り | 複数スライド | PowerPointのスライド数に対応するMarpスライドが `---` で区切られる |
| タイトル | タイトルプレースホルダー、タイトル相当テキスト | 対象テキストがスライド単位 `.drawio.svg` 内のテキスト要素として出力される |
| 本文 | テキストボックス、通常段落 | 本文テキストがスライド単位 `.drawio.svg` 内に出力され、欠落しない |
| 箇条書き | 箇条書き、番号付きリスト、階層付きリスト | 箇条書き、番号付きリスト、階層がスライド単位 `.drawio.svg` 内で視認できる |
| 改行 | テキストボックス内改行、表セル内改行 | 改行がスライド単位 `.drawio.svg` 内のテキストまたは表相当要素で保持される |
| 表 | PowerPoint表、セル結合を含む表 | スライド単位 `.drawio.svg` 内に表相当要素として出力され、表内テキストが含まれる |
| 画像 | PowerPoint内画像、SVG画像 | 画像がスライド単位 `.drawio.svg` 内に埋め込まれ、個別画像ファイルは出力されない |
| 背景画像 | スライド背景画像 | 背景画像がスライド単位 `.drawio.svg` 内に埋め込まれ、個別背景画像ファイルは出力されない |
| 図形テキスト | 図形、テキストボックス | 図形内文字列がスライド単位 `.drawio.svg` 内に出力される。難しい場合は警告として扱う |
| スライドSVG | 各PowerPointスライド | スライド数と同数の `slide/slideNNN.drawio.svg` が出力され、Markdownから参照される |
| draw.io候補 | 図形、線、矢印を含む図解 | 図解がスライド単位 `.drawio.svg` 内に含まれる |
| draw.io直接参照禁止 | draw.io図解 | Markdown内に `.drawio)` の直接画像参照がなく、`.drawio.svg` を参照する |
| 発表者ノート | notesSlideを含むPowerPoint | ノート本文がMarp presenter notes用のHTMLコメントとして出力される |
| 出力先 | 右クリック選択したPowerPointファイル | 入力PowerPointと同一フォルダに入力PowerPointファイル名フォルダが作成され、その配下に `.md` と `slide/` が出力される |
| ファイル名 | 日本語、空白、括弧を含むPowerPointファイル名 | 入力拡張子 `.pptx` を除いた同名フォルダ、その配下の `.md` が出力される |
| 複数スライド | 1〜100スライド | スライド順が維持され、先頭スライドから末尾スライドまでMarkdownに出力される |
| フォルダ一括 | 1〜100ファイル | 入力PowerPoint件数と同数のMarkdownが各入力ファイル名フォルダに出力される |
| 処理時間 | 通常サイズの単一ファイル | 3秒以内にMarkdown、slideまで完了する |
| エラー | Officeとして開けないファイル、パスワード付きファイル | Markdownとslideを出力せず、エラー終了する |
| 警告 | 画像参照切れ、未対応図形、動画、音声 | 変換可能な要素を含むMarkdownを出力し、対象要素は警告として扱う |

## 4. テストケース

| No | 区分 | 入力条件 | 確認内容 | 期待結果 |
|---:|---|---|---|---|
| UT-F04-001 | 基本 | `sample.pptx` | Marp Markdown生成 | `sample/sample.md` が1件出力され、UTF-8で読み込める |
| UT-F04-002 | Marpヘッダー | 任意の `.pptx` | front matter | Markdown先頭に `---`、`marp: true`、`theme:`、`paginate:` が出力される |
| UT-F04-003 | 中間JSON | `sample.pptx` | 中間ファイル出力有無 | 入力ファイル名由来の `.json` ファイルと `poc*_intermediate.json` が出力されない |
| UT-F04-004 | スライド数最小 | 1スライド | スライド生成 | Markdownが1件出力され、`slide/slide001.drawio.svg` への参照が1件含まれる |
| UT-F04-005 | スライド数複数 | 10スライド | スライド区切りと順序 | スライド区切り `---` が必要数出力され、`slide/slide001.drawio.svg` から `slide/slide010.drawio.svg` まで順に参照される |
| UT-F04-006 | スライド数上限 | 100スライド | 処理継続 | Markdownが1件出力され、先頭スライドと末尾スライドの `.drawio.svg` 参照が含まれる |
| UT-F04-007 | タイトル | タイトル付きスライド3枚 | タイトル変換 | 各タイトルが対応するスライド単位 `.drawio.svg` 内のテキスト要素として出力される |
| UT-F04-008 | タイトルなし | タイトルプレースホルダーがないスライド | スライド参照 | タイトルがなくても対応するスライド単位 `.drawio.svg` が出力され、Markdownから参照される |
| UT-F04-009 | 本文 | 通常テキストボックス3件 | 本文出力 | 入力テキスト3件が対応するスライド単位 `.drawio.svg` 内に存在する |
| UT-F04-010 | 箇条書き | 箇条書き3件 | list変換 | 各項目テキストが対応するスライド単位 `.drawio.svg` 内で箇条書きとして視認できる |
| UT-F04-011 | 番号付きリスト | 番号付きリスト3件 | list変換 | 各項目テキストが対応するスライド単位 `.drawio.svg` 内で番号付きリストとして視認できる |
| UT-F04-012 | 階層リスト | 2階層以上の箇条書き | インデント変換 | 子項目が対応するスライド単位 `.drawio.svg` 内で親項目より深い階層として視認できる |
| UT-F04-013 | 改行 | テキストボックス内改行を含むPowerPoint | 改行保持 | 改行前後の文字列が対応するスライド単位 `.drawio.svg` 内で欠落なく出力される |
| UT-F04-014 | 表 | 3行×3列のPowerPoint表1件 | 表変換 | 対応するスライド単位 `.drawio.svg` 内に表相当要素が出力され、表内セル値がすべて含まれる |
| UT-F04-015 | 表セル改行 | セル内改行を含むPowerPoint表1件 | セル内改行変換 | 対象セルの改行が対応するスライド単位 `.drawio.svg` 内の表相当要素で保持される |
| UT-F04-016 | セル結合表 | 行結合または列結合を含むPowerPoint表1件 | 結合セル変換 | 対応するスライド単位 `.drawio.svg` 内で結合セルが近似表現され、セル値が欠落しない |
| UT-F04-017 | 画像1件 | 画像1件を含むPowerPoint | 画像埋め込み | スライド単位 `.drawio.svg` が出力され、個別画像ファイルは出力されない |
| UT-F04-018 | 画像複数 | 画像3件を含むPowerPoint | 画像埋め込み | スライド単位 `.drawio.svg` 内に画像が埋め込まれ、Markdown内の参照は `.drawio.svg` のみである |
| UT-F04-019 | SVG画像 | SVG画像1件を含むPowerPoint | SVG埋め込み | スライド単位 `.drawio.svg` が出力され、個別 `slideNNN-imageNNN.svg` は出力されない |
| UT-F04-020 | 画像相対パス | 画像1件を含むPowerPoint | 参照形式 | Markdown内のスライドSVG参照が `slide/slideNNN.drawio.svg` を参照し、`data:image/` を含まない |
| UT-F04-021 | 背景画像 | 背景画像付きスライド1枚 | 背景埋め込み | 背景画像がスライド単位 `.drawio.svg` に含まれ、個別背景画像ファイルは出力されない |
| UT-F04-022 | 図形テキスト | テキスト入り図形1件 | テキスト抽出 | 図形内文字列が対応するスライド単位 `.drawio.svg` 内に出力される |
| UT-F04-023 | draw.io候補 | 図形、テキスト、線、矢印を含む図解 | draw.io SVG出力 | スライド数と同数の `slide/slideNNN.drawio.svg` が出力される |
| UT-F04-024 | draw.io表示参照 | draw.io候補を含むPowerPoint | Marp表示用参照 | Markdown内に `slide/slideNNN.drawio.svg` を参照するスライドSVG参照が出力される |
| UT-F04-025 | draw.io直接参照禁止 | draw.io候補を含むPowerPoint | 参照禁止確認 | Markdown内に `](slide/*.drawio)` 相当の直接参照が出力されない |
| UT-F04-026 | 図形変換フォールバック | `.drawio.svg` 内で近似できない複雑図形 | フォールバック | 近似変換、テキスト抽出、警告のいずれかで扱われる |
| UT-F04-027 | SmartArt | SmartArtを含むPowerPoint | 未対応要素 | Markdownが出力され、SmartArtは近似変換、テキスト抽出、警告のいずれかで扱われる |
| UT-F04-028 | グラフ | 棒、折れ線、円などのチャートを含むPowerPoint | チャート処理 | `.drawio.svg` 内の画像、データ表相当、警告のいずれかで扱われ、本文全体の変換は継続する |
| UT-F04-029 | 発表者ノート | notesSlideを含むPowerPoint | ノート出力 | Markdown内に `<!--` とノート本文が出力される |
| UT-F04-030 | 複合 | タイトル、本文、リスト、表、画像、図形、ノートを各1件以上含むPowerPoint | 総合変換 | Markdown内にMarpヘッダー、スライド数分の `.drawio.svg` 参照、ノートが存在する |
| UT-F04-031 | 出力先 | `/path/to/sample.pptx` | ファイル名フォルダ出力 | `/path/to/sample/sample.md` と `/path/to/sample/slide/` が出力され、図解がある場合は `slide/slideNNN.drawio.svg` が出力される |
| UT-F04-032 | ファイル名 | 日本語、空白、括弧を含むPowerPointファイル名 | 出力ファイル名 | 入力拡張子 `.pptx` を除いた同名フォルダ、その配下の `.md`、`slide/` が出力される |
| UT-F04-033 | サンプルデータ青 | `testData/青 シンプル ビジネス 企画書 営業 プレゼンテーション.pptx` | 実データ変換 | 同名フォルダ配下にMarp Markdownとslideが出力され、UTF-8で読み込める |
| UT-F04-034 | サンプルデータ白 | `testData/白　ビジネス　プロジェクト進捗報告書　会議のプレゼンテーション.pptx` | 実データ変換 | 同名フォルダ配下にMarp Markdownとslideが出力され、UTF-8で読み込める |
| UT-F04-035 | フォルダ1件 | フォルダ内にPowerPoint 1件 | 一括変換 | Markdownが1件出力される |
| UT-F04-036 | フォルダ複数 | フォルダ内にPowerPoint 10件 | 一括変換 | Markdownが10件出力される |
| UT-F04-037 | フォルダ上限 | フォルダ内にPowerPoint 100件 | 一括変換 | Markdownが100件出力され、各入力ファイルと同名のフォルダと `.md` が存在する |
| UT-F04-038 | 処理時間 | 通常サイズの単一PowerPoint | 実行時間 | 3秒以内にMarkdown、slideまで完了する |
| UT-F04-039 | 処理時間超過 | 画像多数または100スライド相当のPowerPoint | 実行時間と結果 | 処理時間、対象ファイル名、結果ステータスがログに残り、処理が完了または警告終了する |
| UT-F04-040 | 不正ファイル | Officeとして開けないファイル | エラー処理 | Markdownとslideを出力せず、エラー終了する |
| UT-F04-041 | パスワード保護 | 暗号化または保護されたPowerPoint | エラー処理 | Markdownとslideを出力せず、エラー終了する |
| UT-F04-042 | 必須XMLなし | `ppt/presentation.xml` がないpptx | エラー処理 | Markdownとslideを出力せず、エラー終了する |
| UT-F04-043 | slide参照切れ | presentationにはslide参照があるがslide XMLがないpptx | 警告処理 | 参照切れスライドをスキップまたはエラー扱いにし、変換可能なスライドは処理される |
| UT-F04-044 | 画像参照切れ | relsに存在する画像参照先がないpptx | 警告処理 | 参照切れ画像を `.drawio.svg` 内に埋め込まず、変換可能なスライドはMarkdownから参照される |
| UT-F04-045 | notesなし | 発表者ノートがないPowerPoint | ノートなし扱い | Markdownが出力され、空のHTMLコメントを不要に出力しない |
| UT-F04-046 | themeなし | `ppt/theme/theme*.xml` がないpptx | デフォルトテーマ | Markdownが出力され、Marp標準テーマで表示できる |
| UT-F04-047 | 動画音声 | 動画または音声を含むPowerPoint | 未対応メディア | 動画、音声ファイルをMarkdownに埋め込まず、警告として扱う |
| UT-F04-048 | Marp CLI確認 | 画像を含む変換後Markdown | HTML変換 | `npx @marp-team/marp-cli@latest sample.md -o sample.html --allow-local-files` が成功し、HTMLが出力される |

## 5. 性能目標

| 項目 | 目標 |
|---|---|
| 通常サイズの単一PowerPoint | 3秒以内 |
| スライド数 | 100スライドまで対応 |
| 埋め込み画像数 | 100画像まで対応 |
| 表数 | 100表まで対応 |
| draw.io候補数 | 100図解まで対応 |
| フォルダ一括 | 100ファイルまで対応 |

3秒目標は通常サイズの単一ファイルを対象とする。
100スライド、100画像、100表、100図解、100ファイルのケースでは、3秒を超えるケースとして処理完了可否とログ出力を確認する。

## 6. 合格基準

- 正常系の全ケースでMarp Markdownが出力され、UTF-8で読み込めること。
- Markdown先頭にMarp用front matterが出力されること。
- 中間JSONファイルが出力されないこと。
- Markdown内のスライド参照数と `slide/` 配下の `.drawio.svg` 数が入力スライド数と一致すること。
- Markdown内のスライド参照は相対パスであり、`data:image/` を含まないこと。
- PowerPointのスライド順がMarp Markdown上で維持されること。
- 表ケースでは、スライド単位 `.drawio.svg` 内の表相当要素として表内テキストが欠落なく出力されること。
- draw.io候補ケースでは、Marp表示用 `slide/slideNNN.drawio.svg` がスライド単位で出力されること。
- Markdown内では `.drawio` を直接参照せず、`.drawio.svg` を参照すること。
- 発表者ノートケースでは、ノート本文がMarp presenter notes用のHTMLコメントとして出力されること。
- 通常サイズの単一ファイルは3秒以内にMarkdown出力まで完了すること。
- Marp CLI確認ケースでは、`--allow-local-files` 付きHTML変換が成功すること。
- エラー系では、Markdownとslideを出力せずエラー終了すること。
- 警告系では、対象要素を除外またはフォールバックし、変換可能な要素を含むMarkdownが出力されること。
