const path = require("node:path");
const { downloadAndUnzipVSCode, runTests } = require("@vscode/test-electron");

async function resolveVSCodeExecutablePath() {
  return downloadAndUnzipVSCode();
}

async function main() {
  delete process.env.ELECTRON_RUN_AS_NODE;

  const extensionDevelopmentPath = path.resolve(__dirname, "..");
  const extensionTestsPath = path.resolve(__dirname, "vscode", "f01-extension.test.js");
  const vscodeExecutablePath = await resolveVSCodeExecutablePath();

  await runTests({
    extensionDevelopmentPath,
    extensionTestsPath,
    vscodeExecutablePath
  });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
