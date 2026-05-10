import * as path from "path";
import * as vscode from "vscode";
import { convertDocxToMarkdown } from "./converters/docxToMarkdown";
import { convertXlsxToDrawio } from "./converters/xlsxToDrawio";
import { registerF01XlsxToDrawioTests } from "./testController";

const output = vscode.window.createOutputChannel("office to mdrow");

type ConvertKind = "excel-drawio" | "word-md";

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(output);
  registerF01XlsxToDrawioTests(context);
  context.subscriptions.push(
    vscode.commands.registerCommand("officeToMdrow.convertExcelToDrawio", (uri?: vscode.Uri) =>
      convertSelectedFile(uri, "excel-drawio")
    ),
    vscode.commands.registerCommand("officeToMdrow.convertWordToMarkdown", (uri?: vscode.Uri) =>
      convertSelectedFile(uri, "word-md")
    ),
    vscode.commands.registerCommand("officeToMdrow.convertOfficeFile", (uri?: vscode.Uri) =>
      convertSelectedFile(uri)
    )
  );
}

export function deactivate() {
  output.dispose();
}

async function convertSelectedFile(
  uri: vscode.Uri | undefined,
  forcedKind?: ConvertKind
) {
  const target = uri ?? vscode.window.activeTextEditor?.document.uri;
  if (!target || target.scheme !== "file") {
    vscode.window.showErrorMessage("変換対象のOfficeファイルを選択してください。");
    return;
  }

  const kind = forcedKind ?? kindFromPath(target.fsPath);
  if (!kind) {
    vscode.window.showErrorMessage(".xlsx または .docx を選択してください。");
    return;
  }

  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `変換中: ${path.basename(target.fsPath)}`,
      cancellable: false
    },
    async () => {
      const generatedPath = await runConverter(kind, target.fsPath);
      vscode.window.showInformationMessage(`変換完了: ${path.basename(generatedPath)}`);
      await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(generatedPath));
    }
  );
}

function kindFromPath(filePath: string): ConvertKind | undefined {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".xlsx") {
    return "excel-drawio";
  }
  if (ext === ".docx") {
    return "word-md";
  }
  return undefined;
}

async function runConverter(kind: ConvertKind, sourcePath: string): Promise<string> {
  output.show(true);
  output.appendLine(`> office-to-mdrow ${kind} ${sourcePath}`);
  try {
    const generatedPath = kind === "excel-drawio"
      ? await convertXlsxToDrawio(sourcePath)
      : await convertDocxToMarkdown(sourcePath);
    output.appendLine(`Generated: ${generatedPath}`);
    return generatedPath;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    output.appendLine(message);
    vscode.window.showErrorMessage(`変換に失敗しました: ${message}`);
    throw error;
  }
}
