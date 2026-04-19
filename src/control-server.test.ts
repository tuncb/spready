import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import { SpreadyControlClient } from "./control-client";
import { SpreadyControlServer } from "./control-server";
import { WorkbookController } from "./workbook-controller";

test("SpreadyControlServer exposes formula-aware reads over TCP", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    await client.applyTransaction({
      operations: [
        {
          startColumn: 0,
          startRow: 0,
          type: "setRange",
          values: [["4", "5", "=A1+B1"]],
        },
      ],
    });

    const methods = await client.call<string[]>("listMethods");
    const displayRange = await client.getSheetDisplayRange({
      columnCount: 3,
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });
    const cellData = await client.getCellData({
      columnIndex: 2,
      rowIndex: 0,
    });

    assert.ok(methods.includes("getCellData"));
    assert.ok(methods.includes("getSheetDisplayRange"));
    assert.deepEqual(displayRange.values, [["4", "5", "9"]]);
    assert.deepEqual(cellData, {
      columnIndex: 2,
      display: "9",
      input: "=A1+B1",
      isFormula: true,
      rowIndex: 0,
      sheetId: displayRange.sheetId,
      sheetName: displayRange.sheetName,
    });
  } finally {
    await client.close();
    await server.stop();
  }
});

test("SpreadyControlServer saves and opens native workbook files over TCP", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);
  const tempDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-tcp-"));

  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    await client.applyTransaction({
      operations: [
        {
          startColumn: 0,
          startRow: 0,
          type: "setRange",
          values: [["4", "5", "=A1+B1"]],
        },
      ],
    });

    const filePath = path.join(tempDirectory, "numbers.spready");
    const saveResult = await client.saveWorkbookFile({ filePath });

    assert.equal(saveResult.changed, true);
    assert.equal(saveResult.summary.documentFilePath, filePath);
    assert.equal(saveResult.summary.hasUnsavedChanges, false);
    assert.ok((await client.call<string[]>("listMethods")).includes("openWorkbookFile"));
    assert.ok((await client.call<string[]>("listMethods")).includes("saveWorkbookFile"));

    await client.applyTransaction({
      operations: [
        {
          columnIndex: 0,
          rowIndex: 0,
          type: "setCell",
          value: "99",
        },
      ],
    });

    await assert.rejects(
      () => client.openWorkbookFile({ filePath }),
      /discardUnsavedChanges: true/,
    );

    const openResult = await client.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath,
    });
    const displayRange = await client.getSheetDisplayRange({
      columnCount: 3,
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });

    assert.equal(openResult.summary.documentFilePath, filePath);
    assert.equal(openResult.summary.hasUnsavedChanges, false);
    assert.deepEqual(displayRange.values, [["4", "5", "9"]]);
  } finally {
    await client.close();
    await server.stop();
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("SpreadyControlServer creates a new workbook over TCP with unsaved-change guard", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    await client.applyTransaction({
      operations: [
        {
          columnIndex: 0,
          rowIndex: 0,
          type: "setCell",
          value: "draft",
        },
      ],
    });

    await assert.rejects(
      () => client.createNewWorkbook(),
      /discardUnsavedChanges: true/,
    );

    const result = await client.createNewWorkbook({
      discardUnsavedChanges: true,
    });

    assert.equal(result.changed, true);
    assert.equal(result.summary.documentFilePath, undefined);
    assert.equal(result.summary.hasUnsavedChanges, false);
    assert.equal(
      (
        await client.getCellData({
          columnIndex: 0,
          rowIndex: 0,
        })
      ).input,
      "",
    );
  } finally {
    await client.close();
    await server.stop();
  }
});
