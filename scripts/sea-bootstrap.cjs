const path = require("node:path");
const { createRequire } = require("node:module");

const executablePath = process.execPath;
const baseDirectory = path.dirname(executablePath);
const appPath = path.join(baseDirectory, "dist", "index.js");
const originalArgv = process.argv.slice();
const userArgs =
  path.resolve(originalArgv[1] ?? "") === path.resolve(executablePath)
    ? originalArgv.slice(2)
    : originalArgv.slice(1);

process.argv = [executablePath, appPath, ...userArgs];
createRequire(appPath)(appPath);
