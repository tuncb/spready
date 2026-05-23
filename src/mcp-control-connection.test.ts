import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { test } from "node:test";

import { McpControlConnection } from "./mcp-control-connection";
import type { SpreadyControlClient } from "./control-client";
import { SpreadyControlServer } from "./control-server";
import { WorkbookController } from "./workbook-controller";

const visibleAppStatus = {
  focusedWindowCount: 1,
  frontendVisible: true,
  visibleWindowCount: 1,
  windowCount: 1,
};

const hiddenAppStatus = {
  focusedWindowCount: 0,
  frontendVisible: false,
  visibleWindowCount: 0,
  windowCount: 1,
};

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

test("McpControlConnection reports refused TCP connections as actionable status", async () => {
  const connection = new McpControlConnection(
    {
      host: "127.0.0.1",
      openApp: false,
      openAppTimeoutMs: 1000,
      port: 45731,
    },
    {
      createClient: () =>
        ({
          connect: async () => {
            throw Object.assign(new Error("connect ECONNREFUSED 127.0.0.1:45731"), {
              code: "ECONNREFUSED",
            });
          },
        }) as unknown as SpreadyControlClient,
    },
  );

  await assert.rejects(
    () => connection.connectToExisting(),
    /Open Spready at this address through MCP or manually/,
  );
  assert.deepEqual(connection.getStatus(), {
    connected: false,
    lastError:
      "Could not connect to the Spready TCP control server at tcp://127.0.0.1:45731. Open Spready at this address through MCP or manually.",
    state: "disconnected",
  });
});

test("McpControlConnection openAppAndConnect reuses an existing server before launching", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0, {
    getAppStatus: () => visibleAppStatus,
    showApp: () => visibleAppStatus,
  });

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
    assert.deepEqual(connection.getStatus(), {
      appStatus: visibleAppStatus,
      connected: true,
      state: "connected",
      target: {
        host: controlInfo.host,
        port: controlInfo.port,
        source: "argv",
      },
    });
  } finally {
    await server.stop();
  }
});

test("McpControlConnection openAppAndConnect fails when TCP is reachable but frontend is hidden", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0, {
    getAppStatus: () => hiddenAppStatus,
    showApp: () => hiddenAppStatus,
  });

  await server.start();

  const controlInfo = server.getInfo();
  const connection = new McpControlConnection(
    {
      host: controlInfo.host,
      openApp: false,
      openAppTimeoutMs: 1,
      port: controlInfo.port,
    },
    {
      openAppAndWaitForControlTarget: async () => {
        throw new Error("launch should not mask a hidden connected frontend");
      },
    },
  );

  try {
    await assert.rejects(
      () => connection.openAppAndConnect(),
      /no visible frontend window was reported/,
    );
    assert.equal(connection.getStatus().connected, false);
    assert.match(connection.getStatus().lastError ?? "", /no visible frontend window was reported/);
    assert.deepEqual(connection.getStatus().target, {
      host: controlInfo.host,
      port: controlInfo.port,
      source: "argv",
    });
    assert.deepEqual(connection.getStatus().state, "disconnected");
    assert.throws(() => connection.requireConnectedClient(), /Call open_spready_app first/);
  } finally {
    await server.stop();
  }
});

test("McpControlConnection openAppAndConnect shows an already-connected app", async () => {
  class FakeClient extends EventEmitter {
    status = hiddenAppStatus;

    async connect() {
      return undefined;
    }

    async close() {
      return undefined;
    }

    async showApp() {
      this.status = visibleAppStatus;

      return this.status;
    }

    async getAppStatus() {
      return this.status;
    }
  }

  const target = {
    host: "127.0.0.1",
    port: 45731,
    source: "argv" as const,
  };
  const connection = new McpControlConnection(
    {
      host: target.host,
      openApp: false,
      openAppTimeoutMs: 1000,
      port: target.port,
    },
    {
      createClient: () => new FakeClient() as unknown as SpreadyControlClient,
    },
  );

  await connection.connectToExisting();
  const result = await connection.openAppAndConnect();

  assert.deepEqual(result, {
    launched: false,
    target,
  });
  assert.deepEqual(connection.getStatus(), {
    appStatus: visibleAppStatus,
    connected: true,
    state: "connected",
    target,
  });
});

test("McpControlConnection launchAppAndConnect uses the explicit launch path", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0, {
    getAppStatus: () => visibleAppStatus,
    showApp: () => visibleAppStatus,
  });

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
    assert.deepEqual(connection.getStatus().appStatus, visibleAppStatus);
  } finally {
    await server.stop();
  }
});
