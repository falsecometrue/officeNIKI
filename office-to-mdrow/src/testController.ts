import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import AdmZip = require("adm-zip");
import * as vscode from "vscode";
import { convertDocxToMarkdown } from "./converters/docxToMarkdown";
import { XMLParser } from "fast-xml-parser";
import { convertXlsxToDrawio } from "./converters/xlsxToDrawio";

type TestCase = {
  id: string;
  label: string;
  fixture: string;
  feature: "F01" | "F02";
  run: (context: ConvertedFixture) => void;
};

type ConvertedFixture = {
  sourcePath: string;
  outputPath: string;
  xml: string;
  markdown: string;
  resources: string[];
  parsed: any;
};

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true
});

export function registerOfficeToMdrowTests(context: vscode.ExtensionContext): void {
  const controller = vscode.tests.createTestController(
    "officeToMdrowTests",
    "office-to-mdraw"
  );
  context.subscriptions.push(controller);

  const f01TestDataDir = unitTestDataDir(context, "UT_F01_excel→drowio変換機能");
  const f02TestDataDir = unitTestDataDir(context, "UT_F02_word→md変換機能");

  const f01Suite = controller.createTestItem(
    "UT_F01_excel→drowio変換機能",
    "UT_F01 Excel→Draw.io変換機能"
  );
  controller.items.add(f01Suite);

  const f02Suite = controller.createTestItem(
    "UT_F02_word→md変換機能",
    "UT_F02 Word→Markdown変換機能"
  );
  controller.items.add(f02Suite);

  const cases = [...f01TestCases(), ...f02TestCases()];
  const caseByItemId = new Map<string, TestCase>();
  for (const testCase of cases) {
    const item = controller.createTestItem(testCase.id, `${testCase.id}: ${testCase.label}`);
    item.canResolveChildren = false;
    if (testCase.feature === "F01") {
      f01Suite.children.add(item);
    } else {
      f02Suite.children.add(item);
    }
    caseByItemId.set(item.id, testCase);
  }

  controller.createRunProfile(
    "Run",
    vscode.TestRunProfileKind.Run,
    async (request, token) => {
      const run = controller.createTestRun(request);
      const targets = collectRequestedCases(request, [f01Suite, f02Suite], caseByItemId);

      for (const target of targets) {
        if (token.isCancellationRequested) {
          run.skipped(target.item);
          continue;
        }

        run.started(target.item);
        const startedAt = Date.now();
        try {
          const testDataDir = target.testCase.feature === "F01" ? f01TestDataDir : f02TestDataDir;
          if (!fs.existsSync(testDataDir)) {
            throw new Error(`testDataが見つかりません: ${testDataDir}`);
          }
          const converted = target.testCase.feature === "F01"
            ? await convertXlsxFixture(testDataDir, target.testCase.fixture)
            : await convertDocxFixture(testDataDir, target.testCase.fixture);
          target.testCase.run(converted);
          run.passed(target.item, Date.now() - startedAt);
        } catch (error) {
          run.failed(target.item, new vscode.TestMessage(errorMessage(error)), Date.now() - startedAt);
        }
      }

      run.end();
    },
    true
  );
}

function collectRequestedCases(
  request: vscode.TestRunRequest,
  suites: vscode.TestItem[],
  caseByItemId: Map<string, TestCase>
): Array<{ item: vscode.TestItem; testCase: TestCase }> {
  const excludes = new Set((request.exclude || []).map((item) => item.id));
  const collected: Array<{ item: vscode.TestItem; testCase: TestCase }> = [];

  const visit = (item: vscode.TestItem) => {
    if (excludes.has(item.id)) {
      return;
    }
    const testCase = caseByItemId.get(item.id);
    if (testCase) {
      collected.push({ item, testCase });
      return;
    }
    item.children.forEach(visit);
  };

  if (request.include?.length) {
    for (const item of request.include) {
      visit(item);
    }
  } else {
    for (const suite of suites) {
      visit(suite);
    }
  }

  return collected;
}

async function convertXlsxFixture(testDataDir: string, fixtureName: string): Promise<ConvertedFixture> {
  const sourcePath = copyTestFile(testDataDir, fixtureName);
  const outputPath = await convertXlsxToDrawio(sourcePath);
  const xml = fs.readFileSync(outputPath, "utf8");
  return {
    sourcePath,
    outputPath,
    xml,
    markdown: "",
    resources: [],
    parsed: parser.parse(xml)
  };
}

async function convertDocxFixture(testDataDir: string, fixtureName: string): Promise<ConvertedFixture> {
  const sourcePath = copyTestFile(testDataDir, fixtureName);
  const outputPath = await convertDocxToMarkdown(sourcePath);
  const markdown = fs.readFileSync(outputPath, "utf8");
  const resourceDir = path.join(path.dirname(sourcePath), "resources");
  const resources = fs.existsSync(resourceDir)
    ? fs.readdirSync(resourceDir).filter((entry) => fs.statSync(path.join(resourceDir, entry)).isFile())
    : [];
  return {
    sourcePath,
    outputPath,
    xml: "",
    markdown,
    resources,
    parsed: undefined
  };
}

function copyTestFile(testDataDir: string, fixtureName: string): string {
  const source = path.join(testDataDir, fixtureName);
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "office-to-mdraw-test-ui-"));
  const copied = path.join(tempDir, fixtureName);
  fs.copyFileSync(source, copied);
  return copied;
}

function f01TestCases(): TestCase[] {
  return [
    {
      id: "UT-F01-001",
      label: "Draw.io XMLを生成し、XML parseできる",
      fixture: "UT-F01-001_012.xlsx",
      feature: "F01",
      run: ({ outputPath, parsed }) => {
        assertEqual(path.extname(outputPath), ".drawio", "拡張子が.drawioであること");
        assertEqual(fs.existsSync(outputPath), true, "Draw.io XMLが出力されること");
        assertEqual(parsed.mxfile?.host, "app.diagrams.net", "mxfile hostがapp.diagrams.netであること");
      }
    },
    {
      id: "UT-F01-007",
      label: "図形がvertexとして出力される",
      fixture: "UT-F01-007.xlsx",
      feature: "F01",
      run: ({ xml }) => {
        const shapeVertices = xml.match(/<mxCell id="shape-[^"]+"[^>]*vertex="1"/g) || [];
        const shapeValues = xml.match(/<mxCell id="shape-[^"]+"[^>]*value="[^"]+"/g) || [];
        assertOk(shapeVertices.length >= 3, `図形vertex数が3件以上であること。実際: ${shapeVertices.length}`);
        assertOk(shapeValues.length >= 1, "テキスト付き図形が1件以上出力されること");
      }
    },
    {
      id: "UT-F01-008",
      label: "矢印がedgeとして出力され、終端矢印を持つ",
      fixture: "UT-F01-007_008.xlsx",
      feature: "F01",
      run: ({ xml }) => {
        const edges = xml.match(/<mxCell id="edge-[^"]+"[^>]*edge="1"/g) || [];
        assertOk(edges.length >= 1, `edgeが1件以上であること。実際: ${edges.length}`);
        assertOk(xml.includes("endArrow=classic"), "endArrow=classicが含まれること");
      }
    },
    {
      id: "UT-F01-009",
      label: "画像がDraw.io XMLへbase64埋め込みされる",
      fixture: "UT-F01-009.xlsx",
      feature: "F01",
      run: ({ xml }) => {
        const imageShapes = xml.match(/shape=image/g) || [];
        const embeddedImages = xml.match(/data:image\//g) || [];
        assertOk(imageShapes.length >= 2, `画像shape数が2件以上であること。実際: ${imageShapes.length}`);
        assertOk(embeddedImages.length >= 2, `base64画像数が2件以上であること。実際: ${embeddedImages.length}`);
      }
    },
    {
      id: "UT-F01-012",
      label: "入力シートと同数のdiagramを出力し、diagram名がシート名と一致する",
      fixture: "UT-F01-001_012.xlsx",
      feature: "F01",
      run: ({ sourcePath, parsed }) => {
        assertDeepEqual(diagramNames(parsed), workbookSheetNames(sourcePath), "diagram名が入力シート名と一致すること");
      }
    }
  ];
}

function f02TestCases(): TestCase[] {
  return [
    {
      id: "UT-F02-001",
      label: "Markdownを生成し、UTF-8で読み込める",
      fixture: "UT-F02-001.docx",
      feature: "F02",
      run: ({ outputPath, markdown }) => {
        assertEqual(path.extname(outputPath), ".md", "拡張子が.mdであること");
        assertEqual(fs.existsSync(outputPath), true, "Markdownが出力されること");
        assertOk(markdown.length > 0, "Markdown本文が空でないこと");
        assertOk(markdown.includes("新見積管理システム"), "本文テキストが含まれること");
      }
    },
    {
      id: "UT-F02-003",
      label: "見出しをMarkdown見出しとして出力する",
      fixture: "UT-F02-003_004.docx",
      feature: "F02",
      run: ({ markdown }) => {
        assertMatch(markdown, /^# .+/m, "# 見出しが出力されること");
        assertMatch(markdown, /^## .+/m, "## 見出しが出力されること");
        assertMatch(markdown, /^### .+/m, "### 見出しが出力されること");
        assertMatch(markdown, /^#### .+/m, "#### 見出しが出力されること");
      }
    },
    {
      id: "UT-F02-004",
      label: "通常段落を本文として出力し、段落間に空行を入れる",
      fixture: "UT-F02-003_004.docx",
      feature: "F02",
      run: ({ markdown }) => {
        assertOk(markdown.includes("Update/create Word and PowerPoint content"), "通常段落が含まれること");
        assertOk(markdown.includes("To use this document:"), "別の通常段落が含まれること");
        assertMatch(markdown, /\S+\n\n\S+/, "段落間に空行が含まれること");
      }
    },
    {
      id: "UT-F02-012",
      label: "複数画像をresourcesへ出力し、Markdownから相対参照する",
      fixture: "UT-F02-012.docx",
      feature: "F02",
      run: ({ markdown, resources }) => {
        const imageLinks = markdown.match(/!\[[^\]]*]\(resources\/[^)]+\)/g) || [];
        assertEqual(resources.length, 3, "resources配下の画像数が3件であること");
        assertOk(imageLinks.length >= 3, `Markdown画像リンク数が3件以上であること。実際: ${imageLinks.length}`);
        assertOk(!markdown.includes("data:image/"), "data URI画像を含まないこと");
      }
    },
    {
      id: "UT-F02-018",
      label: "Mermaid候補または図形由来テキストをMarkdownへ出力する",
      fixture: "UT-F02-018.docx",
      feature: "F02",
      run: ({ markdown }) => {
        assertOk(
          markdown.includes("```mermaid") || markdown.includes("> 図形テキスト:"),
          "Mermaidコードブロックまたは図形テキストが出力されること"
        );
        assertOk(markdown.includes("開始"), "開始ラベルが含まれること");
        assertOk(markdown.includes("終了"), "終了ラベルが含まれること");
      }
    }
  ];
}

function workbookSheetNames(workbookPath: string): string[] {
  const zip = new AdmZip(workbookPath);
  const workbook = parser.parse(zip.readAsText("xl/workbook.xml"));
  return asArray(workbook.workbook?.sheets?.sheet).map((sheet: any) => String(sheet.name));
}

function diagramNames(parsedDrawio: any): string[] {
  return asArray(parsedDrawio.mxfile?.diagram).map((diagram: any) => String(diagram.name));
}

function asArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function assertOk(condition: boolean, message: string): void {
  if (!condition) {
    throw new Error(message);
  }
}

function assertEqual(actual: unknown, expected: unknown, message: string): void {
  if (actual !== expected) {
    throw new Error(`${message}\nexpected: ${String(expected)}\nactual: ${String(actual)}`);
  }
}

function assertDeepEqual(actual: unknown, expected: unknown, message: string): void {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${message}\nexpected: ${JSON.stringify(expected)}\nactual: ${JSON.stringify(actual)}`);
  }
}

function assertMatch(actual: string, pattern: RegExp, message: string): void {
  if (!pattern.test(actual)) {
    throw new Error(message);
  }
}

function unitTestDataDir(context: vscode.ExtensionContext, featureDirName: string): string {
  return path.resolve(
    context.extensionPath,
    "..",
    "doc",
    "30.developAndTest",
    "01.unitTest",
    featureDirName,
    "testData"
  );
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
