const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const projectRoot = path.resolve(__dirname, "..");
const repoRoot = path.resolve(projectRoot, "..");

function unitTestDataDir(featureDirName) {
  return path.join(
    repoRoot,
    "doc",
    "30.developAndTest",
    "01.unitTest",
    featureDirName,
    "testData"
  );
}

function copyFixture(testDataDir, fixtureName, tempPrefix) {
  const source = path.join(testDataDir, fixtureName);
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), tempPrefix));
  const copied = path.join(tempDir, fixtureName);
  fs.copyFileSync(source, copied);
  return copied;
}

function asArray(value) {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

module.exports = {
  asArray,
  copyFixture,
  projectRoot,
  repoRoot,
  unitTestDataDir
};
