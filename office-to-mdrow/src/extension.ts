import * as path from "path";
import { spawn } from "child_process";
import * as vscode from "vscode";

const output = vscode.window.createOutputChannel("office to mdrow");

type ConvertKind = "excel-drawio" | "word-md";

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(output);
  context.subscriptions.push(
    vscode.commands.registerCommand("officeToMdrow.convertExcelToDrawio", (uri?: vscode.Uri) =>
      convertSelectedFile(context, uri, "excel-drawio")
    ),
    vscode.commands.registerCommand("officeToMdrow.convertWordToMarkdown", (uri?: vscode.Uri) =>
      convertSelectedFile(context, uri, "word-md")
    ),
    vscode.commands.registerCommand("officeToMdrow.convertOfficeFile", (uri?: vscode.Uri) =>
      convertSelectedFile(context, uri)
    )
  );
}

export function deactivate() {
  output.dispose();
}

async function convertSelectedFile(
  context: vscode.ExtensionContext,
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
      const generatedPath = await runConverter(context, kind, target.fsPath);
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

function runConverter(context: vscode.ExtensionContext, kind: ConvertKind, sourcePath: string): Promise<string> {
  const config = vscode.workspace.getConfiguration("officeToMdrow");
  const pythonPath = config.get<string>("pythonPath") || "python3";
  const converter = context.asAbsolutePath(path.join("scripts", "convert.py"));

  output.show(true);
  output.appendLine(`> ${pythonPath} ${converter} ${kind} ${sourcePath}`);

  return new Promise((resolve, reject) => {
    const child = spawn(pythonPath, [converter, kind, sourcePath], {
      cwd: path.dirname(sourcePath)
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (data: Buffer) => {
      const text = data.toString();
      stdout += text;
      output.append(text);
    });

    child.stderr.on("data", (data: Buffer) => {
      const text = data.toString();
      stderr += text;
      output.append(text);
    });

    child.on("error", (error) => {
      vscode.window.showErrorMessage(`変換に失敗しました: ${error.message}`);
      reject(error);
    });

    child.on("close", (code) => {
      if (code === 0) {
        const generated = stdout.trim().split(/\r?\n/).pop();
        resolve(generated || expectedOutputPath(kind, sourcePath));
        return;
      }

      const message = stderr.trim() || stdout.trim() || `exit code ${code}`;
      vscode.window.showErrorMessage(`変換に失敗しました: ${message}`);
      reject(new Error(message));
    });
  });
}

function expectedOutputPath(kind: ConvertKind, sourcePath: string): string {
  const parsed = path.parse(sourcePath);
  const ext = kind === "excel-drawio" ? ".drawio" : ".md";
  return path.join(parsed.dir, `${parsed.name}${ext}`);
}
