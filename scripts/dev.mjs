import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import path from "node:path";

const require = createRequire(import.meta.url);

const viteCli = path.join(path.dirname(require.resolve("vite/package.json")), "bin", "vite.js");
const powerAppsCli = path.join(
  path.dirname(require.resolve("@microsoft/power-apps-cli/package.json")),
  "dist",
  "Bin.js"
);

const children = [
  spawn(process.execPath, [viteCli, "--host", "0.0.0.0", "--port", "3000"], {
    stdio: "inherit",
    env: process.env,
  }),
  spawn(
    process.execPath,
    [powerAppsCli, "run", "--local-app-url", "http://localhost:3000", "--port", "8080"],
    {
      stdio: "inherit",
      env: process.env,
    }
  ),
];

let shuttingDown = false;

function shutdown(exitCode = 0) {
  if (shuttingDown) {
    return;
  }

  shuttingDown = true;

  for (const child of children) {
    if (!child.killed) {
      child.kill("SIGINT");
    }
  }

  setTimeout(() => process.exit(exitCode), 200);
}

for (const child of children) {
  child.on("exit", (code) => {
    shutdown(code ?? 0);
  });
}

process.on("SIGINT", () => shutdown(0));
process.on("SIGTERM", () => shutdown(0));
