import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import AdmZip = require("adm-zip");
import * as vscode from "vscode";
import { XMLParser } from "fast-xml-parser";
import { convertXlsxToDrawio } from "./converters/xlsxToDrawio";

type TestCase = {
  id: string;
  label: string;
  fixture: string;
  run: (context: ConvertedFixture) => void;
};

type ConvertedFixture = {
  workbookPath: string;
  drawioPath: string;
  xml: string;
  parsed: any;
};

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true
});

export function registerF01XlsxToDrawioTests(context: vscode.ExtensionContext): void {
  const controller = vscode.tests.createTestController(
    "officeToMdrowF01Tests",
    "office-to-mdrow F01"
  );
  context.subscriptions.push(controller);

  const testDataDir = path.resolve(
    context.extensionPath,
    "..",
    "doc",
    "30.developAndTest",
    "01.unitTest",
    "UT_F01_excel→drowio変換機能",
    "testData"
  );

  const suite = controller.createTestItem(
    "UT_F01_excel→drowio変換機能",
    "UT_F01 Excel→Draw.io変換機能"
  );
  controller.items.add(suite);

  const cases = f01TestCases();
  const caseByItemId = new Map<string, TestCase>();
  for (const testCase of cases) {
    const item = controller.createTestItem(testCase.id, `${testCase.id}: ${testCase.label}`);
    item.canResolveChildren = false;
    suite.children.add(item);
    caseByItemId.set(item.id, testCase);
  }

  controller.createRunProfile(
    "Run",
    vscode.TestRunProfileKind.Run,
    async (request, token) => {
      const run = controller.createTestRun(request);
      const targets = collectRequestedCases(request, suite, caseByItemId);

      if (!fs.existsSync(testDataDir)) {
        for (const item of targets.map((target) => target.item)) {
          run.errored(item, new vscode.TestMessage(`testDataが見つかりません: ${testDataDir}`));
        }
        run.end();
        return;
      }

      for (const target of targets) {
        if (token.isCancellationRequested) {
          run.skipped(target.item);
          continue;
        }

        run.started(target.item);
        const startedAt = Date.now();
        try {
          const converted = await convertFixture(testDataDir, target.testCase.fixture);
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
  suite: vscode.TestItem,
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
    visit(suite);
  }

  return collected;
}

async function convertFixture(testDataDir: string, fixtureName: string): Promise<ConvertedFixture> {
  const workbookPath = copyTestWorkbook(testDataDir, fixtureName);
  const drawioPath = await convertXlsxToDrawio(workbookPath);
  const xml = fs.readFileSync(drawioPath, "utf8");
  return {
    workbookPath,
    drawioPath,
    xml,
    parsed: parser.parse(xml)
  };
}

function copyTestWorkbook(testDataDir: string, fixtureName: string): string {
  const source = path.join(testDataDir, fixtureName);
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "office-to-mdrow-test-ui-"));
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
      run: ({ drawioPath, parsed }) => {
        assertEqual(path.extname(drawioPath), ".drawio", "拡張子が.drawioであること");
        assertEqual(fs.existsSync(drawioPath), true, "Draw.io XMLが出力されること");
        assertEqual(parsed.mxfile?.host, "app.diagrams.net", "mxfile hostがapp.diagrams.netであること");
      }
    },
    {
      id: "UT-F01-007",
      label: "図形がvertexとして出力される",
      fixture: "UT-F01-007.xlsx",
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
      run: ({ workbookPath, parsed }) => {
        assertDeepEqual(diagramNames(parsed), workbookSheetNames(workbookPath), "diagram名が入力シート名と一致すること");
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

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
