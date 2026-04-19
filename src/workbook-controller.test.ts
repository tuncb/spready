import assert from "node:assert/strict";
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

test("WorkbookController supports raw-vs-display range copy plus explicit paste and clear helpers", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [["1", "2", "=A1+B1"]],
      },
    ],
  });

  const rawCopy = controller.copyRange({
    columnCount: 3,
    mode: "raw",
    rowCount: 1,
    startColumn: 0,
    startRow: 0,
  });
  const displayCopy = controller.copyRange({
    columnCount: 3,
    mode: "display",
    rowCount: 1,
    startColumn: 0,
    startRow: 0,
  });

  assert.equal(rawCopy.text, "1\t2\t=A1+B1");
  assert.deepEqual(rawCopy.values, [["1", "2", "=A1+B1"]]);
  assert.equal(displayCopy.text, "1\t2\t3");
  assert.deepEqual(displayCopy.values, [["1", "2", "3"]]);

  controller.pasteRange({
    startColumn: 0,
    startRow: 1,
    text: displayCopy.text,
  });

  assert.deepEqual(
    controller.getSheetRange({
      columnCount: 3,
      rowCount: 2,
      startColumn: 0,
      startRow: 0,
    }).values,
    [
      ["1", "2", "=A1+B1"],
      ["1", "2", "3"],
    ],
  );

  controller.clearRange({
    columnCount: 2,
    rowCount: 1,
    startColumn: 1,
    startRow: 1,
  });

  assert.deepEqual(
    controller.getSheetRange({
      columnCount: 3,
      rowCount: 2,
      startColumn: 0,
      startRow: 0,
    }).values,
    [
      ["1", "2", "=A1+B1"],
      ["1", "", ""],
    ],
  );
});
