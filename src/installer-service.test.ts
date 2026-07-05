import assert from "node:assert/strict";
import { promises as fs, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import {
  buildDetachedPowerShellStartArguments,
  buildFileAssociationRegistryEntries,
  buildFileOpenCommand,
  buildFinishUpdateScript,
  buildInstallerPowerShellScript,
  buildPowerShellFileArguments,
  buildStartMenuShortcutDetails,
  buildWaitForInstallProcessesExitScript,
  buildUninstallScript,
  buildWaitForProcessExitScript,
  extractSha256FromDigest,
  extractSha256FromText,
  getDefaultInstallDirectory,
  getInstalledExecutablePath,
  getInstallerPowerShellLogPath,
  getStartMenuShortcutPath,
  InstallerService,
  type InstallerCommandRunner,
  type InstallerShortcutDetails,
  parseRegistryDefaultString,
  parseLatestReleaseResponse,
  parseVersionTag,
  runWithAsarFilesystemDisabled,
  selectSha256Asset,
  selectWindowsReleaseAsset,
  UPDATE_RELEASE_URL_ENV_VAR,
  isVersionNewer,
} from "./installer-service";

test("parseVersionTag accepts semantic release tags", () => {
  assert.deepEqual(parseVersionTag("v1.2.3"), {
    major: 1,
    minor: 2,
    patch: 3,
  });
  assert.deepEqual(parseVersionTag("1.2.3-beta.1"), {
    major: 1,
    minor: 2,
    patch: 3,
  });
  assert.equal(parseVersionTag("release-1.2.3"), null);
});

test("isVersionNewer compares semantic versions", () => {
  assert.equal(
    isVersionNewer({ major: 1, minor: 3, patch: 0 }, { major: 1, minor: 2, patch: 9 }),
    true,
  );
  assert.equal(
    isVersionNewer({ major: 1, minor: 2, patch: 9 }, { major: 1, minor: 3, patch: 0 }),
    false,
  );
  assert.equal(
    isVersionNewer({ major: 1, minor: 2, patch: 3 }, { major: 1, minor: 2, patch: 3 }),
    false,
  );
});

test("extractSha256 helpers accept digest fields and checksum text", () => {
  const sha = "a".repeat(64);

  assert.equal(extractSha256FromDigest(`sha256:${sha}`), sha);
  assert.equal(extractSha256FromDigest(sha.toUpperCase()), sha);
  assert.equal(extractSha256FromText(`${sha}  spready.zip`), sha);
  assert.equal(extractSha256FromText("not a checksum"), null);
});

test("release parsing and asset selection target Windows bundles", () => {
  const release = parseLatestReleaseResponse(
    JSON.stringify({
      assets: [
        {
          browser_download_url: "https://example.test/spready-linux-x64-1.2.3.zip",
          name: "spready-linux-x64-1.2.3.zip",
        },
        {
          browser_download_url: "https://example.test/spready-windows-x64-1.2.3.zip",
          digest: `sha256:${"b".repeat(64)}`,
          name: "spready-windows-x64-1.2.3.zip",
        },
        {
          browser_download_url: "https://example.test/spready-windows-x64-1.2.3.zip.sha256",
          name: "spready-windows-x64-1.2.3.zip.sha256",
        },
      ],
      html_url: "https://example.test/releases/v1.2.3",
      tag_name: "v1.2.3",
    }),
  );

  const asset = selectWindowsReleaseAsset(release, "x64");

  assert.equal(asset?.name, "spready-windows-x64-1.2.3.zip");
  assert.equal(selectSha256Asset(release, asset?.name ?? "")?.name, `${asset?.name}.sha256`);
  assert.equal(release.html_url, "https://example.test/releases/v1.2.3");
});

test("update checks use the configured release feed URL", async () => {
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-update-feed-"));
  const releaseUrl = "http://127.0.0.1:45678/latest";
  const env = {
    LOCALAPPDATA: tempDirectory,
    [UPDATE_RELEASE_URL_ENV_VAR]: releaseUrl,
  };
  const installDirectory = getDefaultInstallDirectory(env, "win32");
  const executablePath = getInstalledExecutablePath(installDirectory);
  const requestedUrls: string[] = [];

  try {
    await fs.mkdir(installDirectory, { recursive: true });
    await fs.writeFile(executablePath, "");

    const service = new InstallerService({
      arch: "x64",
      commandRunner: async () => {
        throw new Error("not found");
      },
      currentAppDirectory: installDirectory,
      currentExecutablePath: executablePath,
      currentVersion: "1.2.3",
      env,
      fetch: async (input) => {
        requestedUrls.push(String(input));

        return new Response(
          JSON.stringify({
            assets: [
              {
                browser_download_url: "http://127.0.0.1:45678/spready-windows-x64-1.2.4.zip",
                digest: `sha256:${"c".repeat(64)}`,
                name: "spready-windows-x64-1.2.4.zip",
              },
            ],
            html_url: "http://127.0.0.1:45678/releases/v1.2.4",
            tag_name: "v1.2.4",
          }),
          { status: 200 },
        );
      },
      isPackaged: true,
      platform: "win32",
    });

    const result = await service.checkForUpdates();

    assert.deepEqual(requestedUrls, [releaseUrl]);
    assert.equal(result.updateAvailable, true);
    assert.equal(result.assetName, "spready-windows-x64-1.2.4.zip");
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("install paths are derived from Windows local app data", () => {
  const installDirectory = getDefaultInstallDirectory(
    {
      LOCALAPPDATA: "C:\\Users\\person\\AppData\\Local",
    },
    "win32",
  );

  assert.equal(
    installDirectory,
    path.join("C:\\Users\\person\\AppData\\Local", "Programs", "Spready"),
  );
  assert.equal(
    getInstalledExecutablePath(installDirectory),
    path.join(installDirectory, "Spready.exe"),
  );
});

test("start menu shortcut details include a matching icon index", () => {
  const executablePath = path.join(
    "C:\\Users\\person\\AppData\\Local",
    "Programs",
    "Spready",
    "Spready.exe",
  );

  assert.deepEqual(buildStartMenuShortcutDetails(executablePath), {
    cwd: path.dirname(executablePath),
    description: "Start Spready",
    icon: executablePath,
    iconIndex: 0,
    target: executablePath,
  });
});

test("file association registry entries open workbook files with Spready", () => {
  const executablePath = path.join(
    "C:\\Users\\person\\AppData\\Local",
    "Programs",
    "Spready",
    "Spready.exe",
  );

  assert.equal(buildFileOpenCommand(executablePath), `"${executablePath}" "%1"`);
  assert.deepEqual(buildFileAssociationRegistryEntries(executablePath), [
    {
      key: "HKCU\\Software\\Classes\\.spready",
      value: "Spready.Workbook",
    },
    {
      key: "HKCU\\Software\\Classes\\Spready.Workbook",
      value: "Spready Workbook",
    },
    {
      key: "HKCU\\Software\\Classes\\Spready.Workbook\\DefaultIcon",
      value: `"${executablePath}",0`,
    },
    {
      key: "HKCU\\Software\\Classes\\Spready.Workbook\\shell\\open\\command",
      value: `"${executablePath}" "%1"`,
    },
  ]);
});

test("registry default value parsing reads REG_SZ output", () => {
  assert.equal(
    parseRegistryDefaultString(
      [
        "",
        "HKEY_CURRENT_USER\\Software\\Classes\\.spready",
        "    (Default)    REG_SZ    Spready.Workbook",
        "",
      ].join("\r\n"),
    ),
    "Spready.Workbook",
  );
  assert.equal(parseRegistryDefaultString("value not set"), null);
});

test("non-Windows install directory uses the user data area", () => {
  assert.equal(
    getDefaultInstallDirectory({}, "linux"),
    path.join(os.homedir(), ".local", "share", "Spready"),
  );
});

test("applying installer options writes the start menu shortcut", async () => {
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-installer-shortcut-"));
  const env = {
    APPDATA: tempDirectory,
    LOCALAPPDATA: tempDirectory,
  };
  const installDirectory = getDefaultInstallDirectory(env, "win32");
  const executablePath = getInstalledExecutablePath(installDirectory);
  const expectedShortcutPath = getStartMenuShortcutPath(env, "win32");
  let writtenShortcut: InstallerShortcutDetails | null = null;

  try {
    await fs.mkdir(installDirectory, { recursive: true });
    await fs.writeFile(executablePath, "");

    const service = new InstallerService({
      currentAppDirectory: installDirectory,
      currentExecutablePath: executablePath,
      currentVersion: "1.2.3",
      env,
      isPackaged: true,
      platform: "win32",
      writeShortcut: (shortcutPath, operation, shortcut) => {
        assert.equal(shortcutPath, expectedShortcutPath);
        assert.equal(operation, "create");
        writtenShortcut = shortcut;
        writeFileSync(shortcutPath, "shortcut");

        return true;
      },
    });

    const result = await service.applyOptions({
      fileAssociation: false,
      startMenuShortcut: true,
    });

    assert.deepEqual(writtenShortcut, buildStartMenuShortcutDetails(executablePath));
    assert.equal(result.status.options.startMenuShortcut, true);
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("applying installer options writes the file association", async () => {
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-installer-association-"));
  const env = {
    LOCALAPPDATA: tempDirectory,
  };
  const installDirectory = getDefaultInstallDirectory(env, "win32");
  const executablePath = getInstalledExecutablePath(installDirectory);
  const registry = new Map<string, string>();
  const commandRunner: InstallerCommandRunner = async (command, args) => {
    assert.equal(command, "reg.exe");

    const operation = args[0];
    const key = args[1];

    if (operation === "add") {
      const valueIndex = args.indexOf("/d");

      if (args.includes("/ve") && valueIndex >= 0) {
        registry.set(key, args[valueIndex + 1]);
      }

      return { stderr: "", stdout: "" };
    }

    if (operation === "query") {
      const value = registry.get(key);

      if (!value) {
        throw new Error("not found");
      }

      return {
        stderr: "",
        stdout: ["", key, `    (Default)    REG_SZ    ${value}`, ""].join("\r\n"),
      };
    }

    if (operation === "delete") {
      if (args.includes("/ve")) {
        registry.delete(key);
      } else {
        for (const registryKey of [...registry.keys()]) {
          if (registryKey === key || registryKey.startsWith(`${key}\\`)) {
            registry.delete(registryKey);
          }
        }
      }

      return { stderr: "", stdout: "" };
    }

    throw new Error(`Unexpected reg.exe operation ${operation}.`);
  };

  try {
    await fs.mkdir(installDirectory, { recursive: true });
    await fs.writeFile(executablePath, "");

    const service = new InstallerService({
      commandRunner,
      currentAppDirectory: installDirectory,
      currentExecutablePath: executablePath,
      currentVersion: "1.2.3",
      env,
      isPackaged: true,
      platform: "win32",
    });

    const result = await service.applyOptions({
      fileAssociation: true,
      startMenuShortcut: false,
    });

    assert.equal(registry.get("HKCU\\Software\\Classes\\.spready"), "Spready.Workbook");
    assert.equal(
      registry.get("HKCU\\Software\\Classes\\Spready.Workbook\\shell\\open\\command"),
      `"${executablePath}" "%1"`,
    );
    assert.equal(result.status.options.fileAssociation, true);
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("starting uninstall removes integrations and launches observable cleanup", async () => {
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-installer-uninstall-"));
  const env = {
    APPDATA: tempDirectory,
    LOCALAPPDATA: tempDirectory,
  };
  const installDirectory = getDefaultInstallDirectory(env, "win32");
  const executablePath = getInstalledExecutablePath(installDirectory);
  const shortcutPath = getStartMenuShortcutPath(env, "win32");
  const registry = new Map<string, string>([
    ["HKCU\\Software\\Classes\\.spready", "Spready.Workbook"],
    [
      "HKCU\\Software\\Classes\\Spready.Workbook\\shell\\open\\command",
      buildFileOpenCommand(executablePath),
    ],
    ["HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Spready", ""],
  ]);
  const launcherCalls: Array<{ logPath?: string; operationName: string; script: string }> = [];
  let quitRequested = false;
  const commandRunner: InstallerCommandRunner = async (command, args) => {
    assert.equal(command, "reg.exe");

    const operation = args[0];
    const key = args[1];

    if (operation === "query") {
      const value = registry.get(key);

      if (!value) {
        throw new Error("not found");
      }

      return {
        stderr: "",
        stdout: ["", key, `    (Default)    REG_SZ    ${value}`, ""].join("\r\n"),
      };
    }

    if (operation === "delete") {
      if (args.includes("/ve")) {
        registry.delete(key);
      } else {
        for (const registryKey of [...registry.keys()]) {
          if (registryKey === key || registryKey.startsWith(`${key}\\`)) {
            registry.delete(registryKey);
          }
        }
      }

      return { stderr: "", stdout: "" };
    }

    throw new Error(`Unexpected reg.exe operation ${operation}.`);
  };

  try {
    await fs.mkdir(installDirectory, { recursive: true });
    await fs.writeFile(executablePath, "");
    await fs.mkdir(path.dirname(shortcutPath), { recursive: true });
    await fs.writeFile(shortcutPath, "shortcut");

    const service = new InstallerService({
      commandRunner,
      currentAppDirectory: installDirectory,
      currentExecutablePath: executablePath,
      currentVersion: "1.2.3",
      detachedPowerShellRunner: async (operationName, script, options) => {
        assert.equal(quitRequested, false);
        launcherCalls.push({ logPath: options?.logPath, operationName, script });
      },
      env,
      isPackaged: true,
      platform: "win32",
      requestQuit: () => {
        quitRequested = true;
      },
    });

    const result = await service.startUninstall();

    assert.equal(quitRequested, true);
    assert.equal(registry.has("HKCU\\Software\\Classes\\.spready"), false);
    assert.equal(
      registry.has("HKCU\\Software\\Classes\\Spready.Workbook\\shell\\open\\command"),
      false,
    );
    assert.equal(
      registry.has("HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Spready"),
      false,
    );
    assert.equal(await fileExists(shortcutPath), false);
    assert.equal(launcherCalls.length, 1);
    assert.equal(launcherCalls[0].operationName, "Uninstall Spready");
    assert.match(launcherCalls[0].script, /remove install directory/u);
    assert.match(result.logPath ?? "", /spready-uninstall\.log$/u);
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("installer status reports supported installation options", async () => {
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-installer-status-"));

  try {
    const service = new InstallerService({
      currentAppDirectory: tempDirectory,
      currentExecutablePath: path.join(tempDirectory, "Spready.exe"),
      currentVersion: "1.2.3",
      env: {
        LOCALAPPDATA: tempDirectory,
      },
      isPackaged: true,
      platform: "win32",
    });

    assert.deepEqual((await service.getStatus()).options, {
      fileAssociation: false,
      startMenuShortcut: false,
    });
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("runWithAsarFilesystemDisabled restores the previous ASAR filesystem flag", async () => {
  const processWithAsarFlag = process as NodeJS.Process & {
    noAsar?: boolean;
  };
  const originalNoAsar = processWithAsarFlag.noAsar;

  try {
    processWithAsarFlag.noAsar = false;

    const result = await runWithAsarFilesystemDisabled(async () => {
      assert.equal(processWithAsarFlag.noAsar, true);

      return "copied";
    });

    assert.equal(result, "copied");
    assert.equal(processWithAsarFlag.noAsar, false);
  } finally {
    if (originalNoAsar === undefined) {
      Reflect.deleteProperty(processWithAsarFlag, "noAsar");
    } else {
      processWithAsarFlag.noAsar = originalNoAsar;
    }
  }
});

test("runWithAsarFilesystemDisabled restores the ASAR filesystem flag after failure", async () => {
  const processWithAsarFlag = process as NodeJS.Process & {
    noAsar?: boolean;
  };
  const originalNoAsar = processWithAsarFlag.noAsar;

  try {
    Reflect.deleteProperty(processWithAsarFlag, "noAsar");

    await assert.rejects(
      runWithAsarFilesystemDisabled(async () => {
        assert.equal(processWithAsarFlag.noAsar, true);

        throw new Error("copy failed");
      }),
      /copy failed/u,
    );

    assert.equal(processWithAsarFlag.noAsar, undefined);
  } finally {
    if (originalNoAsar === undefined) {
      Reflect.deleteProperty(processWithAsarFlag, "noAsar");
    } else {
      processWithAsarFlag.noAsar = originalNoAsar;
    }
  }
});

test("update script tolerates the app process already being closed", () => {
  const script = buildWaitForProcessExitScript(12345);

  assert.match(script, /Get-Process -Id 12345 -ErrorAction SilentlyContinue/u);
  assert.doesNotMatch(script, /Wait-Process/u);
});

test("installer PowerShell wrapper logs lifecycle details", () => {
  const script = buildInstallerPowerShellScript(
    "Finish Spready update",
    "Invoke-SpreadyInstallerStep 'copy files' { Write-Host 'copying' }",
    {
      logPath: "C:\\Temp\\spready-installation-logs\\update.log",
      terminalBehavior: "keep-open",
    },
  );

  assert.match(
    script,
    /\$spreadyInstallerLogPath = 'C:\\Temp\\spready-installation-logs\\update.log'/u,
  );
  assert.match(script, /Write-Host \$line/u);
  assert.match(script, /Add-Content -LiteralPath \$spreadyInstallerLogPath -Value \$line/u);
  assert.match(script, /for \(\$attempt = 1; \$attempt -le 20; \$attempt \+= 1\)/u);
  assert.match(script, /Start-Sleep -Milliseconds 50/u);
  assert.match(script, /Starting: \$Name/u);
  assert.match(script, /Write-SpreadyInstallerError \$_/u);
  assert.match(script, /PowerShell window left open for review/u);
});

test("detached installer PowerShell logs use a sibling file", () => {
  assert.equal(
    getInstallerPowerShellLogPath(
      "C:\\Temp\\spready-installation-logs\\2026-06-27-spready-uninstall.log",
    ),
    "C:\\Temp\\spready-installation-logs\\2026-06-27-spready-uninstall.powershell.log",
  );
  assert.equal(
    getInstallerPowerShellLogPath(
      "C:\\Temp\\spready-installation-logs\\2026-06-27-spready-uninstall",
    ),
    "C:\\Temp\\spready-installation-logs\\2026-06-27-spready-uninstall.powershell.log",
  );
});

test("PowerShell file arguments keep installer scripts off the command line", () => {
  const scriptPath = "C:\\Temp\\spready-installation-logs\\update.ps1";

  assert.deepEqual(buildPowerShellFileArguments(scriptPath), [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
  ]);
  assert.deepEqual(buildPowerShellFileArguments(scriptPath, { keepOpen: true }), [
    "-NoExit",
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
  ]);
  assert.equal(buildPowerShellFileArguments(scriptPath).includes("-EncodedCommand"), false);
});

test("detached PowerShell start arguments use an independent Windows process handoff", () => {
  const scriptPath = "C:\\Temp\\spready-installation-logs\\update.ps1";

  assert.deepEqual(buildDetachedPowerShellStartArguments(scriptPath), [
    "/d",
    "/s",
    "/c",
    "start",
    '""',
    "/min",
    "/D",
    "C:\\Temp\\spready-installation-logs",
    "powershell.exe",
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
  ]);
});

test("uninstall script logs explicit steps and propagates failures", () => {
  const script = buildUninstallScript(
    12345,
    "C:\\Users\\person\\AppData\\Local\\Programs\\Spready",
  );

  assert.match(script, /Invoke-SpreadyInstallerStep 'wait for Spready to exit'/u);
  assert.match(script, /Invoke-SpreadyInstallerStep 'remove install directory'/u);
  assert.match(script, /for \(\$attempt = 1; \$attempt -le 60; \$attempt \+= 1\)/u);
  assert.match(script, /Remove-Item -LiteralPath/u);
  assert.equal(script.includes("Install directory removal failed on attempt ${attempt}"), true);
  assert.match(script, /Start-Sleep -Milliseconds 500/u);
  assert.doesNotMatch(script, /SilentlyContinue'\n/u);
});

test("install process wait script checks processes launched from install directory", () => {
  const script = buildWaitForInstallProcessesExitScript(
    "C:\\Users\\person\\AppData\\Local\\Programs\\Spready",
    12,
  );

  assert.match(script, /Get-CimInstance Win32_Process/u);
  assert.match(script, /ExecutablePath\.StartsWith\(\$installRoot/u);
  assert.match(script, /OrdinalIgnoreCase/u);
  assert.match(script, /Spready processes are still running from the install directory/u);
  assert.match(script, /AddSeconds\(12\)/u);
});

test("finish update script retries replacement and restores previous install on failure", () => {
  const script = buildFinishUpdateScript({
    installDirectory: "C:\\Users\\person\\AppData\\Local\\Programs\\Spready",
    latestVersion: "1.2.3",
    pid: 12345,
    restart: true,
    stagedDirectory: "C:\\Temp\\spready-update\\stage",
    updateDirectory: "C:\\Temp\\spready-update",
  });

  assert.match(
    script,
    /Invoke-SpreadyInstallerStep 'wait for installed Spready processes to exit'/u,
  );
  assert.match(script, /Get-CimInstance Win32_Process/u);
  assert.match(script, /Invoke-SpreadyUpdateStep 'rename current install'/u);
  assert.match(script, /Invoke-SpreadyUpdateStep 'move staged update'/u);
  assert.match(script, /Invoke-SpreadyUpdateStep 'remove old install backup'/u);
  assert.match(script, /Invoke-SpreadyUpdateStep 'remove temporary update files'/u);
  assert.match(script, /Invoke-SpreadyUpdateStep 'restore previous install'/u);
  assert.match(script, /Write-SpreadyInstallerLog/u);
  assert.match(script, /Restored previous install after update failure/u);
  assert.match(
    script,
    /reg\.exe add 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Spready' \/v DisplayVersion/u,
  );
  assert.match(script, /Start-Process -FilePath/u);
});

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath);

    return true;
  } catch {
    return false;
  }
}
