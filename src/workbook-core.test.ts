import assert from "node:assert/strict";
import { test } from "node:test";

import {
  applyWorkbookTransaction,
  createSheet,
  createWorkbookState,
  getColumnTitle,
  getSheetCsv,
  getSheetRange,
  getSheetUsedRange,
  getWorkbookSummary,
  normalizeSheet,
  parseCsv,
  parseTsv,
  serializeTsv,
  serializeCsv,
  type WorkbookSheet,
  type WorkbookState,
} from "./workbook-core";

function getActiveSheet(state: WorkbookState): WorkbookSheet {
  const activeSheet = state.sheets.find(
    (sheet) => sheet.id === state.activeSheetId,
  );

  assert.ok(activeSheet, "Expected an active sheet in workbook state.");

  return activeSheet;
}

test("createSheet and normalizeSheet build rectangular matrices", () => {
  assert.deepEqual(createSheet(2.9, 0), [[""], [""]]);
  assert.deepEqual(normalizeSheet([["A"], [], ["B", "C"]]), [
    ["A", ""],
    ["", ""],
    ["B", "C"],
  ]);
  assert.deepEqual(normalizeSheet([]), [[""]]);
});

test("parseCsv handles empty input, CRLF rows, quotes, and embedded newlines", () => {
  assert.deepEqual(parseCsv(""), [[""]]);
  assert.deepEqual(
    parseCsv(
      'Name,Note\r\n"Ada","He said ""hi"""\r\n"Linus","line1\nline2"\r\nSolo',
    ),
    [
      ["Name", "Note"],
      ["Ada", 'He said "hi"'],
      ["Linus", "line1\nline2"],
      ["Solo", ""],
    ],
  );
});

test("parseTsv and serializeTsv handle tabs, quotes, and embedded newlines", () => {
  const values = [
    ["Name", "Note"],
    ["Ada", 'tab\t"quote"'],
    ["Linus", "line1\nline2"],
  ];

  const text = serializeTsv(values);

  assert.equal(
    text,
    'Name\tNote\r\nAda\t"tab\t""quote"""\r\nLinus\t"line1\nline2"',
  );
  assert.deepEqual(parseTsv(text), values);
});

test("serializeCsv trims to the used range and escapes special characters", () => {
  const sheet: WorkbookSheet = {
    cells: [
      ["Name", "Note", ""],
      ["Ada", 'comma, "quote"\nline', ""],
      ["", "", ""],
    ],
    id: "sheet-under-test",
    name: "Sheet Under Test",
  };

  assert.equal(
    serializeCsv(sheet),
    'Name,Note\r\nAda,"comma, ""quote""\nline"',
  );
  assert.equal(
    serializeCsv({
      ...sheet,
      cells: [
        ["", ""],
        ["", ""],
      ],
    }),
    "",
  );
});

test("getColumnTitle covers spreadsheet-style alphabet boundaries", () => {
  assert.equal(getColumnTitle(0), "A");
  assert.equal(getColumnTitle(25), "Z");
  assert.equal(getColumnTitle(26), "AA");
  assert.equal(getColumnTitle(51), "AZ");
  assert.equal(getColumnTitle(52), "BA");
  assert.equal(getColumnTitle(701), "ZZ");
  assert.equal(getColumnTitle(702), "AAA");
});

test("getWorkbookSummary, getSheetRange, and getSheetUsedRange reflect workbook contents", () => {
  const initialState = createWorkbookState();
  const nextState = applyWorkbookTransaction(initialState, {
    operations: [
      {
        startColumn: 1,
        startRow: 1,
        type: "setRange",
        values: [
          ["North", "1200"],
          ["South", "980"],
        ],
      },
    ],
  }).state;

  const summary = getWorkbookSummary(nextState);
  const usedRange = getSheetUsedRange(nextState);
  const focusedRange = getSheetRange(nextState, {
    columnCount: 2,
    rowCount: 2,
    startColumn: 1,
    startRow: 1,
  });
  const boundedRange = getSheetRange(nextState, {
    columnCount: 10.5,
    rowCount: 10.5,
    startColumn: 48.8,
    startRow: 198.8,
  });

  assert.equal(summary.version, 1);
  assert.equal(summary.activeSheetId, initialState.activeSheetId);
  assert.equal(summary.activeSheetName, "Sheet 1");
  assert.equal(summary.sheets.length, 1);
  assert.equal(summary.sheets[0].rowCount, 200);
  assert.equal(summary.sheets[0].columnCount, 50);

  assert.deepEqual(usedRange, {
    columnCount: 3,
    rowCount: 3,
    sheetId: nextState.activeSheetId,
    sheetName: "Sheet 1",
    startColumn: 0,
    startRow: 0,
  });
  assert.deepEqual(focusedRange.values, [
    ["North", "1200"],
    ["South", "980"],
  ]);
  assert.equal(boundedRange.startRow, 198);
  assert.equal(boundedRange.startColumn, 48);
  assert.equal(boundedRange.rowCount, 2);
  assert.equal(boundedRange.columnCount, 2);
});

test("applyWorkbookTransaction returns the original state for empty requests and dry-runs changes safely", () => {
  const initialState = createWorkbookState();
  const emptyResult = applyWorkbookTransaction(initialState, {
    operations: [],
  });

  assert.equal(emptyResult.changed, false);
  assert.equal(emptyResult.state, initialState);

  const dryRunResult = applyWorkbookTransaction(initialState, {
    dryRun: true,
    operations: [
      {
        columnIndex: 0,
        rowIndex: 0,
        type: "setCell",
        value: "draft",
      },
    ],
  });

  assert.equal(dryRunResult.changed, true);
  assert.equal(initialState.version, 0);
  assert.equal(getActiveSheet(initialState).cells[0][0], "");
  assert.equal(getActiveSheet(dryRunResult.state).cells[0][0], "draft");
  assert.equal(dryRunResult.state.version, initialState.version);
});

test("applyWorkbookTransaction manages sheet lifecycle operations", () => {
  const initialState = createWorkbookState();
  const defaultSheetId = initialState.activeSheetId;

  const afterAdd = applyWorkbookTransaction(initialState, {
    operations: [
      {
        activate: false,
        columnCount: 2,
        name: "  Alpha  ",
        rowCount: 3,
        sheetId: "sheet-alpha",
        type: "addSheet",
      },
    ],
  }).state;

  const addedSheet = afterAdd.sheets.find(
    (sheet) => sheet.id === "sheet-alpha",
  );

  assert.ok(addedSheet);
  assert.equal(afterAdd.activeSheetId, defaultSheetId);
  assert.equal(addedSheet.name, "Alpha");
  assert.equal(addedSheet.cells.length, 3);
  assert.equal(addedSheet.cells[0].length, 2);

  const afterRenameAndActivate = applyWorkbookTransaction(afterAdd, {
    operations: [
      {
        sheetId: "sheet-alpha",
        type: "setActiveSheet",
      },
      {
        name: "  Budget  ",
        sheetId: "sheet-alpha",
        type: "renameSheet",
      },
    ],
  }).state;

  assert.equal(afterRenameAndActivate.activeSheetId, "sheet-alpha");
  assert.equal(
    afterRenameAndActivate.sheets.find((sheet) => sheet.id === "sheet-alpha")
      ?.name,
    "Budget",
  );

  const afterDelete = applyWorkbookTransaction(afterRenameAndActivate, {
    operations: [
      {
        sheetId: "sheet-alpha",
        type: "deleteSheet",
      },
    ],
  }).state;

  assert.equal(afterDelete.sheets.length, 1);
  assert.equal(afterDelete.activeSheetId, defaultSheetId);

  const singleSheetState = createWorkbookState();

  assert.throws(
    () =>
      applyWorkbookTransaction(singleSheetState, {
        operations: [
          {
            sheetId: singleSheetState.activeSheetId,
            type: "deleteSheet",
          },
        ],
      }),
    /The last sheet cannot be deleted\./,
  );
});

test("applyWorkbookTransaction expands sheets for setCell and setRange writes", () => {
  const nextState = applyWorkbookTransaction(createWorkbookState(), {
    operations: [
      {
        columnIndex: 50,
        rowIndex: 200,
        type: "setCell",
        value: "edge",
      },
      {
        startColumn: 51,
        startRow: 201,
        type: "setRange",
        values: [["A", "B"], ["C"]],
      },
    ],
  }).state;
  const activeSheet = getActiveSheet(nextState);

  assert.equal(activeSheet.cells.length, 203);
  assert.equal(activeSheet.cells[0].length, 53);
  assert.equal(activeSheet.cells[200][50], "edge");
  assert.equal(activeSheet.cells[201][51], "A");
  assert.equal(activeSheet.cells[201][52], "B");
  assert.equal(activeSheet.cells[202][51], "C");
  assert.equal(activeSheet.cells[202][52], "");
});

test("applyWorkbookTransaction supports structural edits and keeps a minimum 1x1 sheet", () => {
  const seededState = applyWorkbookTransaction(createWorkbookState(), {
    operations: [
      {
        rows: [
          ["A", "B"],
          ["C", "D"],
        ],
        type: "replaceSheet",
      },
    ],
  }).state;

  const editedState = applyWorkbookTransaction(seededState, {
    operations: [
      {
        count: 1,
        rowIndex: 1,
        type: "insertRows",
      },
      {
        columnIndex: 1,
        count: 1,
        type: "insertColumns",
      },
      {
        columnIndex: 1,
        rowIndex: 1,
        type: "setCell",
        value: "X",
      },
      {
        columnCount: 1,
        rowCount: 1,
        startColumn: 0,
        startRow: 0,
        type: "clearRange",
      },
    ],
  }).state;

  assert.deepEqual(getActiveSheet(editedState).cells, [
    ["", "", "B"],
    ["", "X", ""],
    ["C", "", "D"],
  ]);

  const collapsedState = applyWorkbookTransaction(editedState, {
    operations: [
      {
        count: 99,
        rowIndex: 0,
        type: "deleteRows",
      },
      {
        columnIndex: 0,
        count: 99,
        type: "deleteColumns",
      },
    ],
  }).state;

  assert.deepEqual(getActiveSheet(collapsedState).cells, [[""]]);
});

test("applyWorkbookTransaction resizes and replaces sheet contents from CSV metadata-aware updates", () => {
  const initialState = applyWorkbookTransaction(createWorkbookState(), {
    operations: [
      {
        rows: [
          ["A", "B"],
          ["C", "D"],
        ],
        type: "replaceSheet",
      },
    ],
  }).state;

  const resizedState = applyWorkbookTransaction(initialState, {
    operations: [
      {
        columnCount: 1,
        rowCount: 1,
        type: "resizeSheet",
      },
    ],
  }).state;

  assert.deepEqual(getActiveSheet(resizedState).cells, [["A"]]);

  const importedState = applyWorkbookTransaction(resizedState, {
    operations: [
      {
        content: "Region,Revenue\r\nNorth,1200",
        name: "Quarterly",
        sourceFilePath: "C:\\data\\quarterly.csv",
        type: "replaceSheetFromCsv",
      },
    ],
  }).state;
  const activeSheet = getActiveSheet(importedState);

  assert.equal(activeSheet.name, "Quarterly");
  assert.equal(activeSheet.sourceFilePath, "C:\\data\\quarterly.csv");
  assert.deepEqual(activeSheet.cells, [
    ["Region", "Revenue"],
    ["North", "1200"],
  ]);
  assert.equal(getSheetCsv(importedState), "Region,Revenue\r\nNorth,1200");
});

test("applyWorkbookTransaction and sheet reads reject invalid requests", () => {
  const initialState = createWorkbookState();

  assert.throws(
    () =>
      getSheetRange(initialState, {
        columnCount: 1,
        rowCount: 1,
        sheetId: "missing-sheet",
        startColumn: 0,
        startRow: 0,
      }),
    /Sheet "missing-sheet" was not found\./,
  );
  assert.throws(
    () =>
      applyWorkbookTransaction(initialState, {
        operations: [
          {
            count: 0,
            rowIndex: 0,
            type: "insertRows",
          },
        ],
      }),
    /Row insert count must be a positive integer\./,
  );
  assert.throws(
    () =>
      applyWorkbookTransaction(initialState, {
        operations: [
          {
            columnIndex: 0,
            rowIndex: -1,
            type: "setCell",
            value: "bad",
          },
        ],
      }),
    /Row index must be a non-negative integer\./,
  );

  const stateWithExtraSheet = applyWorkbookTransaction(initialState, {
    operations: [
      {
        sheetId: "sheet-duplicate",
        type: "addSheet",
      },
    ],
  }).state;

  assert.throws(
    () =>
      applyWorkbookTransaction(stateWithExtraSheet, {
        operations: [
          {
            sheetId: "sheet-duplicate",
            type: "addSheet",
          },
        ],
      }),
    /Sheet "sheet-duplicate" already exists\./,
  );
});
