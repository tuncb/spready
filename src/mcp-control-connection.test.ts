import assert from "node:assert/strict";
import { test } from "node:test";

import { McpControlConnection } from "./mcp-control-connection";
import { SpreadyControlServer } from "./control-server";
import { WorkbookController } from "./workbook-controller";

test("McpControlConnection starts disconnected and reports a clear workbook-tool error", () => {
  const connection = new McpControlConnection({
    openApp: false,
    openAppTimeoutMs: 1000,
  });

  assert.deepEqual(connection.getStatus(), {
    connected: false,
    state: "disconnected",
  });
  assert.throws(() => connection.requireConnectedClient(), /Call open_spready_app first/);
});

test("McpControlConnection connects to an existing TCP control server", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  const connection = new McpControlConnection({
    host: controlInfo.host,
    openApp: false,
    openAppTimeoutMs: 1000,
    port: controlInfo.port,
  });

  try {
    const result = await connection.connectToExisting();

    assert.equal(result.launched, false);
    assert.deepEqual(result.target, {
      host: controlInfo.host,
      port: controlInfo.port,
      source: "argv",
    });
    assert.deepEqual(connection.getStatus(), {
      connected: true,
      state: "connected",
      target: result.target,
    });
    assert.equal(
      (await connection.requireConnectedClient().getWorkbookSummary()).activeSheetName,
      "Sheet 1",
    );
  } finally {
    await server.stop();
  }
});

test("McpControlConnection openAppAndConnect reuses an existing server before launching", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  const connection = new McpControlConnection(
    {
      host: controlInfo.host,
      openApp: false,
      openAppTimeoutMs: 1000,
      port: controlInfo.port,
    },
    {
      openAppAndWaitForControlTarget: async () => {
        throw new Error("launch should not be needed");
      },
    },
  );

  try {
    const result = await connection.openAppAndConnect();

    assert.equal(result.launched, false);
    assert.equal(connection.getStatus().connected, true);
  } finally {
    await server.stop();
  }
});

test("McpControlConnection launchAppAndConnect uses the explicit launch path", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  let launchCount = 0;
  const connection = new McpControlConnection(
    {
      host: controlInfo.host,
      openApp: true,
      openAppTimeoutMs: 1000,
      port: controlInfo.port,
    },
    {
      openAppAndWaitForControlTarget: async () => {
        launchCount += 1;

        return {
          host: controlInfo.host,
          port: controlInfo.port,
          source: "discovery",
        };
      },
    },
  );

  try {
    const result = await connection.launchAppAndConnect();

    assert.equal(result.launched, true);
    assert.equal(launchCount, 1);
    assert.equal(connection.getStatus().connected, true);
  } finally {
    await server.stop();
  }
});
