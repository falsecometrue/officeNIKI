# UT_F04 PowerPoint→Marp変換機能 テスト仕様書

## 1. 目的

PowerPoint→Marp変換機能について、PowerPoint文書の主要要素、Marp Markdown出力、画像外部管理、draw.io表示用ファイル、フォルダ一括変換、処理時間の観点で確認する。
変換結果は、AIがスライド内容を理解しやすい構造を優先し、次に人間がMarkdownとして編集しやすく、Marpでプレビューできることを確認する。

## 2. テスト方針

- 変換対象は `.pptx` とする。
- 出力は入力ファイルと同一フォルダに作成した入力PowerPointファイル名フォルダ内のMarp Markdownとする。
- 中間JSONファイルは出力しない。
- 画像はbase64ではなく、`resources/` フォルダ配下へ外部ファイルとして出力する。
- Markdown本文から画像、SVG、draw.io表示用ファイルは相対パスで参照する。
- PowerPointの各スライドはMarpの1スライドとして出力する。
- スライド区切りはMarpの `---` を使用する。
- 図形はMermaid化、draw.io候補化、画像フォールバック、警告の順に扱う。
- draw.ioは `.drawio` を直接Marpへ埋め込まず、`.drawio.svg` または `.drawio.png` をMarp表示用画像として参照する。
- PowerPoint表示との完全一致は判定対象外とし、本仕様書に記載したMarkdown記法、出力ファイル数、画像ファイル数、リンク数、処理時間を合格基準とする。

## 3. テスト観点

| 観点 | 確認内容 | 期待結果 |
|---|---|---|
| Marpヘッダー | front matter | Markdown先頭に `marp: true`、`theme:`、`paginate:` が出力される |
| スライド区切り | 複数スライド | PowerPointのスライド数に対応するMarpスライドが `---` で区切られる |
| タイトル | タイトルプレースホルダー、タイトル相当テキスト | 対象テキストが `# ` で始まるMarkdown見出しとして出力される |
| 本文 | テキストボックス、通常段落 | 本文テキストがMarkdown本文として出力され、欠落しない |
| 箇条書き | 箇条書き、番号付きリスト、階層付きリスト | 箇条書きは `- `、番号付きリストは `1. `、階層はインデント付きで出力される |
| 改行 | テキストボックス内改行、表セル内改行 | 本文改行はMarkdownで読める改行、表セル内改行は `<br>` として出力される |
| 表 | PowerPoint表、セル結合を含む表 | HTML tableとして出力され、表内テキストが含まれる |
| 画像 | PowerPoint内画像、SVG画像 | 入力画像数と同数の画像ファイルが `resources/` に出力され、Markdown内に同数の画像リンクが出力される |
| 背景画像 | スライド背景画像 | `resources/` に背景画像が出力され、Marpの `![bg]` 構文で参照される |
| 図形テキスト | 図形、テキストボックス | 図形内テキストがMarkdown本文、Mermaid、draw.io、画像フォールバックのいずれかで出力される |
| Mermaid候補 | 単純なフロー図、時系列図 | Markdown内に ` ```mermaid ` と `flowchart` または `sequenceDiagram` が出力される |
| draw.io候補 | 図形、線、矢印を含む図解 | `diagrams/` に `.drawio` と `.drawio.svg` または `.drawio.png` が出力され、Markdownから表示用ファイルを参照する |
| draw.io直接参照禁止 | draw.io図解 | Markdown内に `.drawio)` の直接画像参照がなく、`.drawio.svg` または `.drawio.png` を参照する |
| 発表者ノート | notesSlideを含むPowerPoint | ノート本文がMarp presenter notes用のHTMLコメントとして出力される |
| 出力先 | 右クリック選択したPowerPointファイル | 入力PowerPointと同一フォルダに入力PowerPointファイル名フォルダが作成され、その配下に `.md`、`resources/`、必要に応じて `diagrams/` が出力される |
| ファイル名 | 日本語、空白、括弧を含むPowerPointファイル名 | 入力拡張子 `.pptx` を除いた同名フォルダ、その配下の `.md` が出力される |
| 変換ログ | 警告、未対応要素 | `convert-report.json` に対象ファイル、警告、フォールバック内容が記録される |
| 複数スライド | 1〜100スライド | スライド順が維持され、先頭スライドから末尾スライドまでMarkdownに出力される |
| フォルダ一括 | 1〜100ファイル | 入力PowerPoint件数と同数のMarkdownが各入力ファイル名フォルダに出力される |
| 処理時間 | 通常サイズの単一ファイル | 3秒以内にMarkdown、resources、convert-report出力まで完了する |
| エラー | Officeとして開けないファイル、パスワード付きファイル | Markdownとresourcesを出力せず、エラー終了する |
| 警告 | 画像参照切れ、未対応図形、動画、音声 | 変換可能な要素を含むMarkdownを出力し、対象要素は警告として扱う |

## 4. テストケース

| No | 区分 | 入力条件 | 確認内容 | 期待結果 |
|---:|---|---|---|---|
| UT-F04-001 | 基本 | `sample.pptx` | Marp Markdown生成 | `sample/sample.md` が1件出力され、UTF-8で読み込める |
| UT-F04-002 | Marpヘッダー | 任意の `.pptx` | front matter | Markdown先頭に `---`、`marp: true`、`theme:`、`paginate:` が出力される |
| UT-F04-003 | 中間JSON | `sample.pptx` | 中間ファイル出力有無 | 入力ファイル名由来の `.json` ファイルと `poc*_intermediate.json` が出力されない。ただし `convert-report.json` は出力対象とする |
| UT-F04-004 | スライド数最小 | 1スライド | スライド生成 | Markdownが1件出力され、本文テキストが1件以上含まれる |
| UT-F04-005 | スライド数複数 | 10スライド | スライド区切りと順序 | スライド区切り `---` が必要数出力され、先頭スライドの文字列が前方、末尾スライドの文字列が後方に出力される |
| UT-F04-006 | スライド数上限 | 100スライド | 処理継続 | Markdownが1件出力され、先頭スライドと末尾スライドの文字列が含まれる |
| UT-F04-007 | タイトル | タイトル付きスライド3枚 | 見出し変換 | 各タイトルが `# ` で始まる行として出力される |
| UT-F04-008 | タイトルなし | タイトルプレースホルダーがないスライド | 補助見出し | `# Slide N` などの補助見出し、または本文のみでスライドが欠落なく出力される |
| UT-F04-009 | 本文 | 通常テキストボックス3件 | 本文出力 | 入力テキスト3件がすべてMarkdown内に存在する |
| UT-F04-010 | 箇条書き | 箇条書き3件 | list変換 | `- ` で始まる行が3行出力され、各項目テキストが含まれる |
| UT-F04-011 | 番号付きリスト | 番号付きリスト3件 | list変換 | `1. ` で始まる行が3行出力され、各項目テキストが含まれる |
| UT-F04-012 | 階層リスト | 2階層以上の箇条書き | インデント変換 | 子項目が親項目より深いインデントで出力される |
| UT-F04-013 | 改行 | テキストボックス内改行を含むPowerPoint | 改行保持 | 改行前後の文字列がMarkdown内で欠落なく出力される |
| UT-F04-014 | 表 | 3行×3列のPowerPoint表1件 | HTML table変換 | Markdown内に `<table>`、`<tr>`、`<td>` または `<th>` が出力され、表内セル値がすべて含まれる |
| UT-F04-015 | 表セル改行 | セル内改行を含むPowerPoint表1件 | セル内改行変換 | 対象セルのMarkdown出力に `<br>` が1件以上含まれる |
| UT-F04-016 | セル結合表 | 行結合または列結合を含むPowerPoint表1件 | 結合セル変換 | HTML table内に `rowspan` または `colspan` が出力され、セル値が欠落しない |
| UT-F04-017 | 画像1件 | 画像1件を含むPowerPoint | 画像外部出力 | `resources/` に画像ファイルが1件出力され、Markdown内に `![` で始まる画像リンクが1件出力される |
| UT-F04-018 | 画像複数 | 画像3件を含むPowerPoint | 画像数一致 | `resources/` の画像ファイルが3件、Markdown内の画像リンクが3件出力される |
| UT-F04-019 | SVG画像 | SVG画像1件を含むPowerPoint | SVG外部出力 | `resources/slideNNN-imageNNN.svg` が出力され、Markdownから相対参照される |
| UT-F04-020 | 画像相対パス | 画像1件を含むPowerPoint | 画像参照形式 | Markdown内の画像リンクが `resources/` で始まる相対パスを参照し、`data:image/` を含まない |
| UT-F04-021 | 背景画像 | 背景画像付きスライド1枚 | 背景出力 | `resources/slideNNN-backgroundNNN.*` が出力され、Markdown内に `![bg](` が1件以上出力される |
| UT-F04-022 | 図形テキスト | テキスト入り図形1件 | テキスト抽出 | 図形内文字列がMarkdown本文、Mermaid、draw.io、画像リンクのいずれかに出力される |
| UT-F04-023 | Mermaidフロー図 | 開始、処理、終了、矢印を含む単純フロー図 | Mermaid出力 | Markdown内に ` ```mermaid `、`flowchart`、開始/処理/終了のラベルが各1件以上出力される |
| UT-F04-024 | Mermaidシーケンス図 | 横方向に並ぶ主体と矢印を含む単純図 | Mermaid出力 | Markdown内に ` ```mermaid `、`sequenceDiagram`、主体名またはメッセージが出力される |
| UT-F04-025 | draw.io候補 | 図形、テキスト、線、矢印を含む図解 | draw.io出力 | `diagrams/` に `.drawio` と `.drawio.svg` または `.drawio.png` が出力される |
| UT-F04-026 | draw.io表示参照 | draw.io候補を含むPowerPoint | Marp表示用参照 | Markdown内に `diagrams/slide` を含む画像リンクが出力され、参照先拡張子が `.drawio.svg` または `.drawio.png` である |
| UT-F04-027 | draw.io直接参照禁止 | draw.io候補を含むPowerPoint | 参照禁止確認 | Markdown内に `](diagrams/*.drawio)` 相当の直接参照が出力されない |
| UT-F04-028 | 図形画像フォールバック | Mermaid化、draw.io化できない複雑図形 | フォールバック | 取得可能な場合は画像リンクとして出力され、取得できない場合は `convert-report.json` に警告が出力される |
| UT-F04-029 | SmartArt | SmartArtを含むPowerPoint | 未対応要素 | Markdownが出力され、SmartArtは画像フォールバックまたは警告として扱われる |
| UT-F04-030 | グラフ | 棒、折れ線、円などのチャートを含むPowerPoint | チャート処理 | Mermaid候補、データ表、画像フォールバック、警告のいずれかで扱われ、本文全体の変換は継続する |
| UT-F04-031 | 発表者ノート | notesSlideを含むPowerPoint | ノート出力 | Markdown内に `<!--` とノート本文が出力される |
| UT-F04-032 | 複合 | タイトル、本文、リスト、表、画像、図形、ノートを各1件以上含むPowerPoint | 総合変換 | Markdown内にMarpヘッダー、見出し、本文、list行、HTML table、画像リンク、図形由来テキストまたはMermaid/draw.io、ノートが各1件以上存在する |
| UT-F04-033 | 出力先 | `/path/to/sample.pptx` | ファイル名フォルダ出力 | `/path/to/sample/sample.md`、`/path/to/sample/resources/`、`/path/to/sample/convert-report.json` が出力される |
| UT-F04-034 | ファイル名 | 日本語、空白、括弧を含むPowerPointファイル名 | 出力ファイル名 | 入力拡張子 `.pptx` を除いた同名フォルダ、その配下の `.md`、`resources/`、`convert-report.json` が出力される |
| UT-F04-035 | サンプルデータ青 | `testData/青 シンプル ビジネス 企画書 営業 プレゼンテーション.pptx` | 実データ変換 | 同名フォルダ配下にMarp Markdownとresourcesが出力され、UTF-8で読み込める |
| UT-F04-036 | サンプルデータ白 | `testData/白　ビジネス　プロジェクト進捗報告書　会議のプレゼンテーション.pptx` | 実データ変換 | 同名フォルダ配下にMarp Markdownとresourcesが出力され、UTF-8で読み込める |
| UT-F04-037 | フォルダ1件 | フォルダ内にPowerPoint 1件 | 一括変換 | Markdownが1件出力される |
| UT-F04-038 | フォルダ複数 | フォルダ内にPowerPoint 10件 | 一括変換 | Markdownが10件出力される |
| UT-F04-039 | フォルダ上限 | フォルダ内にPowerPoint 100件 | 一括変換 | Markdownが100件出力され、各入力ファイルと同名のフォルダと `.md` が存在する |
| UT-F04-040 | 処理時間 | 通常サイズの単一PowerPoint | 実行時間 | 3秒以内にMarkdown、resources、convert-report出力まで完了する |
| UT-F04-041 | 処理時間超過 | 画像多数または100スライド相当のPowerPoint | 実行時間と結果 | 処理時間、対象ファイル名、結果ステータスがログまたは `convert-report.json` に残り、処理が完了または警告終了する |
| UT-F04-042 | 不正ファイル | Officeとして開けないファイル | エラー処理 | Markdownとresourcesを出力せず、エラー終了する |
| UT-F04-043 | パスワード保護 | 暗号化または保護されたPowerPoint | エラー処理 | Markdownとresourcesを出力せず、エラー終了する |
| UT-F04-044 | 必須XMLなし | `ppt/presentation.xml` がないpptx | エラー処理 | Markdownとresourcesを出力せず、エラー終了する |
| UT-F04-045 | slide参照切れ | presentationにはslide参照があるがslide XMLがないpptx | 警告処理 | 参照切れスライドをスキップまたはエラー扱いにし、結果が `convert-report.json` に記録される |
| UT-F04-046 | 画像参照切れ | relsに存在する画像参照先がないpptx | 警告処理 | 参照切れ画像リンクを出力せず、他の本文要素を含むMarkdownが出力される |
| UT-F04-047 | notesなし | 発表者ノートがないPowerPoint | ノートなし扱い | Markdownが出力され、空のHTMLコメントを不要に出力しない |
| UT-F04-048 | themeなし | `ppt/theme/theme*.xml` がないpptx | デフォルトテーマ | Markdownが出力され、Marp標準テーマで表示できる |
| UT-F04-049 | 動画音声 | 動画または音声を含むPowerPoint | 未対応メディア | 動画、音声ファイルをMarkdownに埋め込まず、`convert-report.json` に警告が出力される |
| UT-F04-050 | Marp CLI確認 | 画像を含む変換後Markdown | HTML変換 | `npx @marp-team/marp-cli@latest sample.md -o sample.html --allow-local-files` が成功し、HTMLが出力される |

## 5. 性能目標

| 項目 | 目標 |
|---|---|
| 通常サイズの単一PowerPoint | 3秒以内 |
| スライド数 | 100スライドまで対応 |
| 画像数 | 100画像まで対応 |
| 表数 | 100表まで対応 |
| draw.io候補数 | 100図解まで対応 |
| フォルダ一括 | 100ファイルまで対応 |

3秒目標は通常サイズの単一ファイルを対象とする。
100スライド、100画像、100表、100図解、100ファイルのケースでは、3秒を超えるケースとして処理完了可否とログ出力を確認する。

## 6. 合格基準

- 正常系の全ケースでMarp Markdownが出力され、UTF-8で読み込めること。
- Markdown先頭にMarp用front matterが出力されること。
- 中間JSONファイルが出力されないこと。ただし `convert-report.json` は出力対象とする。
- 画像ケースでは、Markdown内の画像リンク数と `resources/` 配下の画像ファイル数が入力画像数と一致すること。
- 画像リンクは相対パスであり、`data:image/` を含まないこと。
- PowerPointのスライド順がMarp Markdown上で維持されること。
- 表ケースでは、HTML tableとして表内テキストが欠落なく出力されること。
- Mermaid候補ケースでは、Mermaidコードブロックと対象図種別のキーワードが出力されること。
- draw.io候補ケースでは、編集用 `.drawio` とMarp表示用 `.drawio.svg` または `.drawio.png` が出力されること。
- Markdown内では `.drawio` を直接参照せず、`.drawio.svg` または `.drawio.png` を参照すること。
- 発表者ノートケースでは、ノート本文がMarp presenter notes用のHTMLコメントとして出力されること。
- 通常サイズの単一ファイルは3秒以内にMarkdown出力まで完了すること。
- Marp CLI確認ケースでは、`--allow-local-files` 付きHTML変換が成功すること。
- エラー系では、Markdownとresourcesを出力せずエラー終了すること。
- 警告系では、対象要素を除外またはフォールバックし、変換可能な要素を含むMarkdownが出力されること。
- 警告、未対応要素、フォールバック内容が `convert-report.json` に記録されること。
