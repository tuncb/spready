import assert from "node:assert/strict";
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
    const rawCopy = await client.copyRange({
      columnCount: 3,
      mode: "raw",
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });
    const displayCopy = await client.copyRange({
      columnCount: 3,
      mode: "display",
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });
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
    assert.ok(methods.includes("copyRange"));
    assert.ok(methods.includes("pasteRange"));
    assert.ok(methods.includes("clearRange"));
    assert.equal(rawCopy.text, "4\t5\t=A1+B1");
    assert.equal(displayCopy.text, "4\t5\t9");
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
