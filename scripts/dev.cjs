const { spawn } = require("node:child_process");
const http = require("node:http");
const path = require("node:path");

const rootDir = path.resolve(__dirname, "..");
const isWindows = process.platform === "win32";
const npmCommand = isWindows ? "npm.cmd" : "npm";
const electronCommand = isWindows ? "npx.cmd" : "npx";
const devServerUrl = "http://127.0.0.1:5173";

let electronProcess;

function waitForServer(url, retries = 60) {
  return new Promise((resolve, reject) => {
    const attempt = (remaining) => {
      http
        .get(url, (response) => {
          response.resume();
          resolve();
        })
        .on("error", () => {
          if (remaining <= 0) {
            reject(new Error("Vite dev server did not start in time."));
            return;
          }

          setTimeout(() => attempt(remaining - 1), 500);
        });
    };

    attempt(retries);
  });
}

function shutdown(code = 0) {
  if (electronProcess && !electronProcess.killed) {
    electronProcess.kill("SIGTERM");
  }

  process.exit(code);
}

const viteProcess = spawn(
  npmCommand,
  ["run", "dev:renderer", "--", "--host", "0.0.0.0", "--strictPort"],
  {
    cwd: rootDir,
    stdio: "inherit",
    shell: false
  }
);

viteProcess.on("exit", (code) => {
  if (!electronProcess) {
    process.exit(code ?? 0);
    return;
  }

  shutdown(code ?? 0);
});

waitForServer(devServerUrl)
  .then(() => {
    electronProcess = spawn(electronCommand, ["electron", "."], {
      cwd: rootDir,
      stdio: "inherit",
      shell: false,
      env: {
        ...process.env,
        VITE_DEV_SERVER_URL: devServerUrl
      }
    });

    electronProcess.on("exit", (code) => {
      if (!viteProcess.killed) {
        viteProcess.kill("SIGTERM");
      }

      process.exit(code ?? 0);
    });
  })
  .catch((error) => {
    console.error(error.message);
    shutdown(1);
  });

process.on("SIGINT", () => shutdown(0));
process.on("SIGTERM", () => shutdown(0));
