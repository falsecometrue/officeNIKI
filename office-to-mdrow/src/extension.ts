import * as fs from "fs";
import * as path from "path";
import * as vscode from "vscode";
import { convertDocxToMarkdown } from "./converters/docxToMarkdown";
import { convertPptxToMarp } from "./converters/pptxToMarp";
import { convertXlsxToDrawio } from "./converters/xlsxToDrawio";
import { convertXlsxToMarkdown } from "./converters/xlsxToMarkdown";
import { registerOfficeToMdrowTests } from "./testController";

const output = vscode.window.createOutputChannel("office to mdraw");

type ConvertKind = "excel-drawio" | "excel-md" | "word-md" | "powerpoint-marp";

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(output);
  registerOfficeToMdrowTests(context);
  context.subscriptions.push(
    vscode.commands.registerCommand("officeToMdrow.convertExcelToDrawio", (uri?: vscode.Uri) =>
      convertSelectedFile(uri, "excel-drawio")
    ),
    vscode.commands.registerCommand("officeToMdrow.convertExcelToMarkdown", (uri?: vscode.Uri) =>
      convertSelectedFile(uri, "excel-md")
    ),
    vscode.commands.registerCommand("officeToMdrow.convertWordToMarkdown", (uri?: vscode.Uri) =>
      convertSelectedFile(uri, "word-md")
    ),
    vscode.commands.registerCommand("officeToMdrow.convertPowerPointToMarp", (uri?: vscode.Uri) =>
      convertSelectedFile(uri, "powerpoint-marp")
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
    vscode.window.showErrorMessage(".xlsx, .docx, .pptx のいずれかを選択してください。");
    return;
  }

  if (!vscode.workspace.isTrusted) {
    vscode.window.showErrorMessage("未信頼ワークスペースではOfficeファイル変換を実行できません。ワークスペースを信頼してから再実行してください。");
    return;
  }

  const shouldContinue = await confirmOverwriteIfNeeded(kind, target.fsPath);
  if (!shouldContinue) {
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
    return "excel-md";
  }
  if (ext === ".docx") {
    return "word-md";
  }
  if (ext === ".pptx") {
    return "powerpoint-marp";
  }
  return undefined;
}

function plannedOutputPath(kind: ConvertKind, sourcePath: string): string {
  const parsed = path.parse(sourcePath);
  if (kind === "excel-drawio") {
    return path.join(parsed.dir, `${parsed.name}.drawio`);
  }
  return path.join(parsed.dir, parsed.name);
}

async function confirmOverwriteIfNeeded(kind: ConvertKind, sourcePath: string): Promise<boolean> {
  const outputPath = plannedOutputPath(kind, sourcePath);
  if (!fs.existsSync(outputPath)) {
    return true;
  }

  const outputName = path.basename(outputPath);
  const choice = await vscode.window.showWarningMessage(
    `既存の変換結果「${outputName}」を上書きします。よろしいですか？`,
    { modal: true },
    "上書きする"
  );
  return choice === "上書きする";
}

function safeErrorMessage(error: unknown, sourcePath: string): string {
  const raw = error instanceof Error ? error.message : String(error);
  const basename = path.basename(sourcePath);
  const dirname = path.dirname(sourcePath);
  return raw
    .split(sourcePath).join(basename)
    .split(dirname).join("[source directory]")
    .slice(0, 500);
}

async function runConverter(kind: ConvertKind, sourcePath: string): Promise<string> {
  output.show(true);
  output.appendLine(`> office-to-mdraw ${kind} ${path.basename(sourcePath)}`);
  try {
    let generatedPath: string;
    if (kind === "excel-drawio") {
      generatedPath = await convertXlsxToDrawio(sourcePath);
    } else if (kind === "excel-md") {
      generatedPath = await convertXlsxToMarkdown(sourcePath);
    } else if (kind === "powerpoint-marp") {
      generatedPath = await convertPptxToMarp(sourcePath);
    } else {
      generatedPath = await convertDocxToMarkdown(sourcePath);
    }
    output.appendLine(`Generated: ${path.basename(generatedPath)}`);
    return generatedPath;
  } catch (error) {
    const message = safeErrorMessage(error, sourcePath);
    output.appendLine(message);
    vscode.window.showErrorMessage(`変換に失敗しました: ${message}`);
    throw error;
  }
}
