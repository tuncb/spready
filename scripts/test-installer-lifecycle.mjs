#!/usr/bin/env node
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import { promises as fs } from "node:fs";
import { createServer } from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const SCRIPT_PATH = fileURLToPath(import.meta.url);
const REPO_ROOT = path.resolve(path.dirname(SCRIPT_PATH), "..");
const APP_EXECUTABLE_NAME = "Spready.exe";
const MARKER_FILE_NAME = "spready-lifecycle-marker.txt";
const DEFAULT_CURRENT_VERSION = "0.0.1";
const DEFAULT_LATEST_VERSION = "0.0.2";
const DEFAULT_HOLD_MS = 5000;
const UPDATE_WAIT_MS = 180000;

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});

async function main() {
  const args = parseArgs(process.argv.slice(2));

  if (args.help) {
    console.log(getHelpText());
    return;
  }

  if (args.child) {
    await runChildOperation(args.child, args);
    return;
  }

  await runLifecycleTest(args);
}

function getHelpText() {
  return [
    "Spready installer lifecycle test",
    "",
    "Usage:",
    "  npm run package",
    "  npm run test:installer-lifecycle -- -- --bundle out/Spready-win32-x64",
    "",
    "Options:",
    "  --bundle DIR             Packaged Spready directory to install as the old version.",
    "                           Defaults to out/Spready-win32-<arch>.",
    "  --update-bundle DIR      Packaged Spready directory to serve as the fake update.",
    "                           Defaults to --bundle.",
    `  --current-version VER    Version reported by the installed test app. Default ${DEFAULT_CURRENT_VERSION}.`,
    `  --latest-version VER     Version served by the fake release feed. Default ${DEFAULT_LATEST_VERSION}.`,
    `  --hold-ms MS             Milliseconds to keep Spready.exe running during update/uninstall. Default ${DEFAULT_HOLD_MS}.`,
    "  --keep-temp              Keep the temporary profile, install, zip, and logs paths.",
    "",
    "The test uses a temp LOCALAPPDATA/APPDATA profile and a no-op registry runner, so it does",
    "not intentionally modify the user's real Spready install or HKCU registry entries.",
  ].join("\n");
}

function parseArgs(argv) {
  const args = {};

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];

    if (token === "--") {
      continue;
    }

    if (token === "--help" || token === "-h") {
      args.help = true;
      continue;
    }

    if (token === "--keep-temp") {
      args.keepTemp = true;
      continue;
    }

    if (!token.startsWith("--")) {
      throw new Error(`Unexpected argument "${token}".`);
    }

    const separatorIndex = token.indexOf("=");
    const name = token.slice(2, separatorIndex === -1 ? undefined : separatorIndex);
    const inlineValue = separatorIndex === -1 ? undefined : token.slice(separatorIndex + 1);
    const value = inlineValue ?? argv[index + 1];

    if (!value || value.startsWith("--")) {
      throw new Error(`Missing value for --${name}.`);
    }

    args[toCamelCase(name)] = value;

    if (inlineValue === undefined) {
      index += 1;
    }
  }

  return args;
}

function toCamelCase(value) {
  return value.replaceAll(/-([a-z])/gu, (_match, letter) => letter.toUpperCase());
}

async function runLifecycleTest(args) {
  if (process.platform !== "win32") {
    throw new Error("Installer lifecycle testing is only supported on Windows.");
  }

  const arch = args.arch ?? (process.arch === "arm64" ? "arm64" : "x64");
  const bundleDirectory = path.resolve(
    REPO_ROOT,
    args.bundle ?? path.join("out", `Spready-win32-${arch}`),
  );
  const updateBundleDirectory = path.resolve(
    REPO_ROOT,
    args.updateBundle ?? path.relative(REPO_ROOT, bundleDirectory),
  );
  const currentVersion = args.currentVersion ?? DEFAULT_CURRENT_VERSION;
  const latestVersion = args.latestVersion ?? DEFAULT_LATEST_VERSION;
  const holdMilliseconds = parsePositiveInteger(args.holdMs ?? `${DEFAULT_HOLD_MS}`, "hold-ms");

  await assertRegularFile(path.join(bundleDirectory, APP_EXECUTABLE_NAME));
  await assertRegularFile(path.join(updateBundleDirectory, APP_EXECUTABLE_NAME));

  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "spready-installer-lifecycle-"));
  const profileDirectory = path.join(tempRoot, "profile");
  const installDirectory = path.join(profileDirectory, "Programs", "Spready");
  const updateStagingDirectory = path.join(tempRoot, "update-bundle");
  const assetName = `spready-windows-${arch}-${latestVersion}.zip`;
  const assetPath = path.join(tempRoot, assetName);

  log(`temp root: ${tempRoot}`);
  log(`bundle: ${bundleDirectory}`);
  log(`update bundle: ${updateBundleDirectory}`);

  try {
    await runHarnessChild("install", {
      arch,
      bundle: bundleDirectory,
      currentVersion,
      profileDir: profileDirectory,
    });
    await fs.writeFile(path.join(installDirectory, MARKER_FILE_NAME), "old\n", "utf8");
    log(`installed old test build at ${installDirectory}`);

    await fs.rm(updateStagingDirectory, { force: true, recursive: true });
    await fs.mkdir(updateStagingDirectory, { recursive: true });
    await fs.cp(updateBundleDirectory, updateStagingDirectory, {
      dereference: false,
      recursive: true,
    });
    await fs.writeFile(
      path.join(updateStagingDirectory, MARKER_FILE_NAME),
      `updated ${latestVersion}\n`,
      "utf8",
    );
    await compressDirectoryContents(updateStagingDirectory, assetPath);

    const assetSha256 = await sha256File(assetPath);
    const releaseServer = await startReleaseServer({
      assetName,
      assetPath,
      assetSha256,
      latestVersion,
    });

    try {
      await runUpdateScenario({
        arch,
        currentVersion,
        holdMilliseconds,
        installDirectory,
        latestVersion,
        profileDirectory,
        releaseUrl: releaseServer.releaseUrl,
      });
    } finally {
      await releaseServer.close();
    }

    await runUninstallScenario({
      arch,
      currentVersion: latestVersion,
      holdMilliseconds,
      installDirectory,
      profileDirectory,
    });

    log("installer lifecycle test passed");
  } finally {
    if (args.keepTemp) {
      log(`kept temp root: ${tempRoot}`);
    } else {
      await fs.rm(tempRoot, { force: true, recursive: true });
    }
  }
}

async function runUpdateScenario(options) {
  log("starting installed Spready.exe before update");
  const appProcess = await launchInstalledApp(options.installDirectory, options.profileDirectory);
  const stopApp = delay(options.holdMilliseconds).then(() => stopProcessTree(appProcess));

  try {
    const updateResult = await runHarnessChild("update", {
      arch: options.arch,
      currentVersion: options.currentVersion,
      installDir: options.installDirectory,
      profileDir: options.profileDirectory,
      releaseUrl: options.releaseUrl,
    });

    await stopApp;

    await waitFor(async () => {
      const marker = await readTextIfExists(path.join(options.installDirectory, MARKER_FILE_NAME));

      return marker === `updated ${options.latestVersion}\n`;
    }, "updated install marker");

    await waitForLogMessage(
      getInstallerPowerShellLogPath(updateResult.logPath),
      "Finished Finish Spready update.",
    );
    await waitFor(async () => {
      const processes = await getProcessesFromInstallDirectory(options.installDirectory);

      return processes.length === 0;
    }, "no Spready processes from the updated install directory");
    await assertNoBackupDirectories(options.installDirectory);

    log(`update succeeded; log: ${updateResult.logPath}`);
  } finally {
    await stopApp;
  }
}

async function runUninstallScenario(options) {
  log("starting installed Spready.exe before uninstall");
  const appProcess = await launchInstalledApp(options.installDirectory, options.profileDirectory);
  const stopApp = delay(options.holdMilliseconds).then(() => stopProcessTree(appProcess));

  try {
    const uninstallResult = await runHarnessChild("uninstall", {
      arch: options.arch,
      currentVersion: options.currentVersion,
      installDir: options.installDirectory,
      profileDir: options.profileDirectory,
    });

    await stopApp;
    await waitFor(async () => !(await pathExists(options.installDirectory)), "install removal");
    await waitForLogMessage(
      getInstallerPowerShellLogPath(uninstallResult.logPath),
      "Finished Uninstall Spready.",
    );
    log(`uninstall succeeded; log: ${uninstallResult.logPath}`);
  } finally {
    await stopApp;
  }
}

async function runChildOperation(mode, args) {
  const installerModule = await import(
    pathToFileURL(path.join(REPO_ROOT, "src", "installer-service.ts")).href
  );
  const {
    InstallerService,
    UPDATE_RELEASE_URL_ENV_VAR,
    getDefaultInstallDirectory,
    getInstalledExecutablePath,
  } = installerModule;
  const profileDirectory = requireArgument(args, "profileDir");
  const childEnv = {
    ...process.env,
    APPDATA: profileDirectory,
    LOCALAPPDATA: profileDirectory,
    ...(args.releaseUrl ? { [UPDATE_RELEASE_URL_ENV_VAR]: args.releaseUrl } : {}),
  };
  const installDirectory = args.installDir ?? getDefaultInstallDirectory(childEnv, "win32");
  const installedExecutablePath = getInstalledExecutablePath(installDirectory);
  const currentVersion = requireArgument(args, "currentVersion");
  let shouldQuit = false;

  const service = new InstallerService({
    arch: args.arch ?? process.arch,
    commandRunner: runNoOpRegistryCommand,
    currentAppDirectory: args.bundle ?? installDirectory,
    currentExecutablePath: args.bundle
      ? path.join(args.bundle, APP_EXECUTABLE_NAME)
      : installedExecutablePath,
    currentVersion,
    env: childEnv,
    isPackaged: true,
    platform: "win32",
    requestQuit: () => {
      shouldQuit = true;
    },
  });

  let result;

  switch (mode) {
    case "install":
      result = await service.installCurrentApp({
        fileAssociation: false,
        startMenuShortcut: false,
      });
      break;
    case "update":
      result = await service.checkForUpdates({
        restart: false,
        startUpdate: true,
      });
      break;
    case "uninstall":
      result = await service.startUninstall();
      break;
    default:
      throw new Error(`Unknown child operation "${mode}".`);
  }

  console.log(`SPREADY_INSTALLER_LIFECYCLE_RESULT ${JSON.stringify(result)}`);

  if (shouldQuit) {
    setTimeout(() => process.exit(0), 25);
  }
}

async function runNoOpRegistryCommand(command, args) {
  if (command.toLowerCase() !== "reg.exe") {
    throw new Error(`Unexpected installer command "${command}".`);
  }

  if (args[0]?.toLowerCase() === "query") {
    throw new Error("not found");
  }

  return {
    stderr: "",
    stdout: "",
  };
}

async function runHarnessChild(mode, options) {
  const childArgs = [
    "--import",
    "tsx",
    SCRIPT_PATH,
    "--child",
    mode,
    "--arch",
    options.arch,
    "--current-version",
    options.currentVersion,
    "--profile-dir",
    options.profileDir,
    ...(options.bundle ? ["--bundle", options.bundle] : []),
    ...(options.installDir ? ["--install-dir", options.installDir] : []),
    ...(options.releaseUrl ? ["--release-url", options.releaseUrl] : []),
  ];
  const result = await runProcess(process.execPath, childArgs, {
    cwd: REPO_ROOT,
    timeoutMilliseconds: 90000,
  });
  const resultLine = result.stdout
    .split(/\r?\n/u)
    .find((line) => line.startsWith("SPREADY_INSTALLER_LIFECYCLE_RESULT "));

  if (!resultLine) {
    throw new Error(
      [
        `Installer child operation "${mode}" did not report a result.`,
        result.stdout.trim(),
        result.stderr.trim(),
      ]
        .filter(Boolean)
        .join("\n"),
    );
  }

  return JSON.parse(resultLine.replace("SPREADY_INSTALLER_LIFECYCLE_RESULT ", ""));
}

async function launchInstalledApp(installDirectory, profileDirectory) {
  const executablePath = path.join(installDirectory, APP_EXECUTABLE_NAME);
  const child = spawn(executablePath, [], {
    cwd: installDirectory,
    env: {
      ...process.env,
      APPDATA: profileDirectory,
      LOCALAPPDATA: profileDirectory,
      SPREADY_CONTROL_PORT: "0",
    },
    stdio: "ignore",
    windowsHide: true,
  });

  await new Promise((resolve, reject) => {
    child.once("spawn", resolve);
    child.once("error", reject);
  });
  await delay(1500);

  if (child.exitCode !== null) {
    throw new Error(`Installed Spready.exe exited early with code ${child.exitCode}.`);
  }

  return child;
}

async function stopProcessTree(child) {
  if (!child.pid || child.exitCode !== null) {
    return;
  }

  await runProcess("taskkill.exe", ["/PID", `${child.pid}`, "/T", "/F"], {
    allowFailure: true,
    timeoutMilliseconds: 30000,
  });
}

async function startReleaseServer(options) {
  const server = createServer(async (request, response) => {
    try {
      const requestUrl = new URL(request.url ?? "/", "http://127.0.0.1");

      if (requestUrl.pathname === "/latest") {
        const body = JSON.stringify({
          assets: [
            {
              browser_download_url: `http://127.0.0.1:${server.address().port}/${options.assetName}`,
              digest: `sha256:${options.assetSha256}`,
              name: options.assetName,
            },
          ],
          html_url: `http://127.0.0.1:${server.address().port}/releases/v${options.latestVersion}`,
          tag_name: `v${options.latestVersion}`,
        });

        response.writeHead(200, {
          "Content-Length": Buffer.byteLength(body),
          "Content-Type": "application/json",
        });
        response.end(body);
        return;
      }

      if (requestUrl.pathname === `/${options.assetName}`) {
        const assetStats = await fs.stat(options.assetPath);

        response.writeHead(200, {
          "Content-Length": assetStats.size,
          "Content-Type": "application/zip",
        });
        createReadStream(options.assetPath).pipe(response);
        return;
      }

      response.writeHead(404);
      response.end("not found");
    } catch (error) {
      response.writeHead(500);
      response.end(error instanceof Error ? error.message : "server error");
    }
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      server.off("error", reject);
      resolve();
    });
  });

  const address = server.address();

  if (!address || typeof address === "string") {
    throw new Error("Fake release server did not bind to a TCP port.");
  }

  const releaseUrl = `http://127.0.0.1:${address.port}/latest`;

  log(`fake release feed: ${releaseUrl}`);

  return {
    close: () =>
      new Promise((resolve, reject) => {
        server.close((error) => {
          if (error) {
            reject(error);
          } else {
            resolve();
          }
        });
      }),
    releaseUrl,
  };
}

async function compressDirectoryContents(sourceDirectory, destinationPath) {
  await fs.rm(destinationPath, { force: true });

  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$items = @(Get-ChildItem -LiteralPath ${quotePowerShellString(sourceDirectory)} -Force)`,
    "if ($items.Count -eq 0) { throw 'Update bundle directory is empty.' }",
    `Compress-Archive -Path $items.FullName -DestinationPath ${quotePowerShellString(
      destinationPath,
    )} -Force`,
  ].join("\n");

  await runProcess(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
    {
      timeoutMilliseconds: 120000,
    },
  );
}

async function getProcessesFromInstallDirectory(installDirectory) {
  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$installRoot = [System.IO.Path]::GetFullPath(${quotePowerShellString(
      installDirectory,
    )}).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar`,
    "$processes = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {",
    "  $_.ExecutablePath -and $_.ExecutablePath.StartsWith($installRoot, [System.StringComparison]::OrdinalIgnoreCase)",
    "} | Select-Object ProcessId, Name, ExecutablePath)",
    "$processes | ConvertTo-Json -Compress",
  ].join("\n");
  const result = await runProcess(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
    {
      timeoutMilliseconds: 30000,
    },
  );
  const output = result.stdout.trim();

  if (!output) {
    return [];
  }

  const parsed = JSON.parse(output);

  return Array.isArray(parsed) ? parsed : [parsed];
}

async function assertNoBackupDirectories(installDirectory) {
  const parentDirectory = path.dirname(installDirectory);
  const installName = path.basename(installDirectory);
  const entries = await fs.readdir(parentDirectory);
  const backupDirectories = entries.filter((entry) => entry.startsWith(`${installName}.old-`));

  if (backupDirectories.length > 0) {
    throw new Error(`Update left backup directories behind: ${backupDirectories.join(", ")}`);
  }
}

async function waitForLogMessage(logPath, message) {
  await waitFor(async () => {
    const log = await readTextIfExists(logPath);

    return log.includes(message);
  }, `log message "${message}"`);
}

async function waitFor(predicate, description, timeoutMilliseconds = UPDATE_WAIT_MS) {
  const deadline = Date.now() + timeoutMilliseconds;

  while (Date.now() < deadline) {
    if (await predicate()) {
      return;
    }

    await delay(500);
  }

  throw new Error(`Timed out waiting for ${description}.`);
}

async function runProcess(command, args, options = {}) {
  const child = spawn(command, args, {
    cwd: options.cwd,
    stdio: ["ignore", "pipe", "pipe"],
    windowsHide: true,
  });
  let stdout = "";
  let stderr = "";
  let timeout;

  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => {
    stdout += chunk;
  });
  child.stderr.on("data", (chunk) => {
    stderr += chunk;
  });

  const exit = await new Promise((resolve, reject) => {
    child.once("error", reject);
    child.once("exit", (code) => {
      resolve(code ?? 0);
    });

    if (options.timeoutMilliseconds) {
      timeout = setTimeout(() => {
        child.kill();
        reject(new Error(`${command} timed out after ${options.timeoutMilliseconds} ms.`));
      }, options.timeoutMilliseconds);
    }
  }).finally(() => {
    if (timeout) {
      clearTimeout(timeout);
    }
  });

  if (exit !== 0 && !options.allowFailure) {
    throw new Error(
      [
        `${command} exited with code ${exit}.`,
        stdout.trim() ? `stdout:\n${stdout.trim()}` : "",
        stderr.trim() ? `stderr:\n${stderr.trim()}` : "",
      ]
        .filter(Boolean)
        .join("\n"),
    );
  }

  return {
    stderr,
    stdout,
  };
}

async function sha256File(filePath) {
  const content = await fs.readFile(filePath);

  return createHash("sha256").update(content).digest("hex");
}

async function assertRegularFile(filePath) {
  try {
    const stats = await fs.stat(filePath);

    if (stats.isFile()) {
      return;
    }
  } catch {
    // Use the common error below.
  }

  throw new Error(`Expected file at ${filePath}. Run npm run package or pass --bundle.`);
}

async function readTextIfExists(filePath) {
  try {
    return await fs.readFile(filePath, "utf8");
  } catch {
    return "";
  }
}

async function pathExists(filePath) {
  try {
    await fs.access(filePath);

    return true;
  } catch {
    return false;
  }
}

function requireArgument(args, name) {
  const value = args[name];

  if (!value) {
    throw new Error(`Missing required argument --${name}.`);
  }

  return value;
}

function parsePositiveInteger(value, name) {
  const parsed = Number.parseInt(value, 10);

  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`--${name} must be a positive integer.`);
  }

  return parsed;
}

function quotePowerShellString(value) {
  return `'${value.replaceAll("'", "''")}'`;
}

function getInstallerPowerShellLogPath(logPath) {
  return logPath.toLowerCase().endsWith(".log")
    ? `${logPath.slice(0, -4)}.powershell.log`
    : `${logPath}.powershell.log`;
}

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function log(message) {
  console.log(`[spready-installer-lifecycle] ${message}`);
}
