import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import {
  getMcpStdioHelpText,
  getPackagedAppExecutablePath,
  parseMcpStartupOptions,
  resolveSpreadyAppLaunchCommand,
  waitForControlTarget,
} from "./mcp-startup";
import { SpreadyControlServer } from "./control-server";
import { WorkbookController } from "./workbook-controller";

test("parseMcpStartupOptions reads open app and control connection flags", () => {
  assert.deepEqual(
    parseMcpStartupOptions([
      "--openApp",
      "--host=127.0.0.1",
      "--port",
      "47500",
      "--appPath",
      "C:\\Spready\\Spready.exe",
      "--openAppTimeoutMs=5000",
    ]),
    {
      appPath: "C:\\Spready\\Spready.exe",
      help: false,
      host: "127.0.0.1",
      openApp: true,
      openAppTimeoutMs: 5000,
      port: 47500,
    },
  );
});

test("parseMcpStartupOptions accepts dashed aliases", () => {
  assert.deepEqual(parseMcpStartupOptions(["--open-app", "--app-path=/tmp/Spready"]), {
    appPath: "/tmp/Spready",
    help: false,
    openApp: true,
    openAppTimeoutMs: 20000,
  });
});

test("parseMcpStartupOptions recognizes help flags", () => {
  assert.deepEqual(parseMcpStartupOptions(["--help"]), {
    help: true,
    openApp: false,
    openAppTimeoutMs: 20000,
  });
  assert.deepEqual(parseMcpStartupOptions(["-h"]), {
    help: true,
    openApp: false,
    openAppTimeoutMs: 20000,
  });
});

test("getMcpStdioHelpText documents startup options", () => {
  const helpText = getMcpStdioHelpText("spready-mcp");

  assert.match(helpText, /Usage: spready-mcp \[options\]/);
  assert.match(helpText, /--help/);
  assert.match(helpText, /--openApp/);
  assert.match(helpText, /--appPath/);
  assert.match(helpText, /--open-app-timeout-ms/);
});

test("getPackagedAppExecutablePath resolves macOS app bundle executables", () => {
  assert.equal(
    getPackagedAppExecutablePath("/Applications/Spready.app", "darwin"),
    path.join("/Applications/Spready.app", "Contents", "MacOS", "Spready"),
  );
  assert.equal(
    getPackagedAppExecutablePath("C:\\Spready\\Spready.exe", "win32"),
    "C:\\Spready\\Spready.exe",
  );
});

test("resolveSpreadyAppLaunchCommand uses an explicit app path and forwards requested port", async () => {
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-mcp-startup-"));
  const executablePath = path.join(
    tempDirectory,
    process.platform === "win32" ? "Spready.exe" : "Spready",
  );

  await fs.writeFile(executablePath, "", "utf8");

  try {
    const command = await resolveSpreadyAppLaunchCommand({
      appPath: executablePath,
      env: {},
      port: 47501,
    });

    assert.equal(command.command, executablePath);
    assert.deepEqual(command.args, []);
    assert.equal(command.env.SPREADY_CONTROL_PORT, "47501");
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("resolveSpreadyAppLaunchCommand uses cmd.exe for Windows npm dev launches", async () => {
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-mcp-startup-"));

  await fs.writeFile(
    path.join(tempDirectory, "package.json"),
    JSON.stringify({ name: "spready" }),
    "utf8",
  );

  try {
    const command = await resolveSpreadyAppLaunchCommand({
      cwd: tempDirectory,
      env: {},
      platform: "win32",
    });

    assert.equal(command.command, process.env.ComSpec ?? "cmd.exe");
    assert.deepEqual(command.args, ["/d", "/s", "/c", "npm.cmd run start"]);
    assert.equal(command.cwd, tempDirectory);
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("waitForControlTarget resolves once the TCP control server accepts connections", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();

  try {
    const target = await waitForControlTarget({
      host: controlInfo.host,
      port: controlInfo.port,
      timeoutMs: 1000,
    });

    assert.deepEqual(target, {
      host: controlInfo.host,
      port: controlInfo.port,
      source: "argv",
    });
  } finally {
    await server.stop();
  }
});
