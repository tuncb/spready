import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import { WorkbookController } from "./workbook-controller";

test("WorkbookController exposes raw range reads separately from display-range and cell-data reads", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["1", "2", "=A1+B1"],
          ["text", "", "=A2+1"],
        ],
      },
    ],
  });

  const rawRange = controller.getSheetRange({
    columnCount: 3,
    rowCount: 2,
    startColumn: 0,
    startRow: 0,
  });
  const displayRange = controller.getSheetDisplayRange({
    columnCount: 3,
    rowCount: 2,
    startColumn: 0,
    startRow: 0,
  });
  const formulaCell = controller.getCellData({
    columnIndex: 2,
    rowIndex: 0,
  });
  const valueErrorCell = controller.getCellData({
    columnIndex: 2,
    rowIndex: 1,
  });

  assert.deepEqual(rawRange.values, [
    ["1", "2", "=A1+B1"],
    ["text", "", "=A2+1"],
  ]);
  assert.deepEqual(displayRange.values, [
    ["1", "2", "3"],
    ["text", "", "#VALUE!"],
  ]);
  assert.deepEqual(formulaCell, {
    columnIndex: 2,
    display: "3",
    input: "=A1+B1",
    isFormula: true,
    rowIndex: 0,
    sheetId: rawRange.sheetId,
    sheetName: rawRange.sheetName,
  });
  assert.deepEqual(valueErrorCell, {
    columnIndex: 2,
    display: "#VALUE!",
    errorCode: "VALUE",
    input: "=A2+1",
    isFormula: true,
    rowIndex: 1,
    sheetId: rawRange.sheetId,
    sheetName: rawRange.sheetName,
  });
  assert.equal(controller.getSummary().hasUnsavedChanges, true);
});

test("WorkbookController keeps CSV export on raw input strings even when formulas are present", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["1", "=A1+2"],
          ["", "=B1*2"],
        ],
      },
    ],
  });

  assert.equal(controller.getSheetCsv(), "1,=A1+2\r\n,=B1*2");
  assert.deepEqual(
    controller.getSheetDisplayRange({
      columnCount: 2,
      rowCount: 2,
      startColumn: 0,
      startRow: 0,
    }).values,
    [
      ["1", "3"],
      ["", "6"],
    ],
  );
});

test("WorkbookController saves and opens native workbook files", async () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["1", "2", "=A1+B1"],
          ["North", "980", ""],
        ],
      },
      {
        sourceFilePath: "C:\\imports\\north.csv",
        type: "setSheetSourceFile",
      },
      {
        activate: true,
        columnCount: 2,
        name: "Notes",
        rowCount: 2,
        type: "addSheet",
      },
      {
        columnIndex: 0,
        rowIndex: 0,
        type: "setCell",
        value: "Saved",
      },
    ],
  });

  const tempDirectory = await fs.mkdtemp(
    path.join(os.tmpdir(), "spready-controller-"),
  );

  try {
    assert.equal(controller.getSummary().hasUnsavedChanges, true);

    const saveResult = await controller.saveWorkbookFile({
      filePath: path.join(tempDirectory, "budget"),
    });

    assert.equal(saveResult.changed, true);
    assert.equal(saveResult.summary.documentFilePath, saveResult.filePath);
    assert.equal(saveResult.summary.hasUnsavedChanges, false);
    assert.match(saveResult.filePath, /\.spready$/);
    assert.match(
      await fs.readFile(saveResult.filePath, "utf8"),
      /"format": "spready-workbook"/,
    );

    controller.applyTransaction({
      operations: [
        {
          columnIndex: 0,
          rowIndex: 0,
          type: "setCell",
          value: "Changed",
        },
      ],
    });

    assert.equal(controller.getSummary().hasUnsavedChanges, true);

    await assert.rejects(
      () =>
        controller.openWorkbookFile({
          filePath: saveResult.filePath,
        }),
      /discardUnsavedChanges: true/,
    );

    const openResult = await controller.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath: saveResult.filePath,
    });

    assert.equal(openResult.changed, true);
    assert.equal(openResult.summary.documentFilePath, saveResult.filePath);
    assert.equal(openResult.summary.hasUnsavedChanges, false);
    assert.equal(
      controller.getSheetDisplayRange({
        columnCount: 3,
        rowCount: 2,
        sheetId: openResult.summary.sheets[0].id,
        startColumn: 0,
        startRow: 0,
      }).values[0][2],
      "3",
    );
    assert.equal(
      controller.getCellData({
        columnIndex: 0,
        rowIndex: 0,
      }).input,
      "Saved",
    );

    const secondSaveResult = await controller.saveWorkbookFile({
      filePath: saveResult.filePath,
    });

    assert.equal(secondSaveResult.changed, false);
    assert.equal(secondSaveResult.summary.hasUnsavedChanges, false);
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});
