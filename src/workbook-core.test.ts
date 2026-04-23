import assert from "node:assert/strict";
import { test } from "node:test";

import {
  adjustWorkbookChartForDeletedColumns,
  adjustWorkbookChartForDeletedRows,
  adjustWorkbookChartForInsertedColumns,
  adjustWorkbookChartForInsertedRows,
  applyWorkbookTransaction,
  createWorkbookChartSummary,
  createSheet,
  createWorkbookState,
  getColumnTitle,
  getWorkbookChartDimensionCount,
  getWorkbookChartStatus,
  getWorkbookChartValidationIssues,
  getSheetColumnCount,
  getSheetCsv,
  getSheetRange,
  getSheetStyleRange,
  getSheetRowCount,
  getSheetUsedRange,
  getWorkbookSummary,
  normalizeSheet,
  parseCsv,
  parseTsv,
  serializeTsv,
  serializeCsv,
  type WorkbookChart,
  type WorkbookChartCartesianSpec,
  type WorkbookChartSheetReference,
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

function createBarChart(
  overrides: Partial<Pick<WorkbookChart, "id" | "layout" | "name" | "sheetId">> = {},
): WorkbookChart & { spec: WorkbookChartCartesianSpec } {
  return {
    id: "chart-1",
    layout: {
      height: 260,
      offsetX: 0,
      offsetY: 0,
      startColumn: 5,
      startRow: 2,
      width: 420,
      zIndex: 0,
    },
    name: "Revenue",
    sheetId: "sheet-1",
    spec: {
      categoryDimension: 0,
      chartType: "bar",
      family: "cartesian",
      source: {
        range: {
          columnCount: 3,
          rowCount: 4,
          sheetId: "sheet-1",
          startColumn: 1,
          startRow: 2,
        },
        seriesLayoutBy: "column",
        sourceHeader: true,
      },
      valueDimensions: [1, 2],
    },
    ...overrides,
  };
}

function getChartSheetReferences(
  workbook: Pick<WorkbookState, "sheets">,
): WorkbookChartSheetReference[] {
  return workbook.sheets.map((sheet) => ({
    columnCount: getSheetColumnCount(sheet),
    id: sheet.id,
    rowCount: getSheetRowCount(sheet),
  }));
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
    cellStyles: {},
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

test("workbook chart helpers validate same-sheet range contracts and summarize status", () => {
  const sheets: WorkbookChartSheetReference[] = [
    {
      columnCount: 10,
      id: "sheet-1",
      rowCount: 12,
    },
  ];
  const validChart = createBarChart();
  const invalidChart: WorkbookChart = {
    ...validChart,
    id: "chart-2",
    sheetId: "sheet-2",
    spec: {
      ...validChart.spec,
      source: {
        ...validChart.spec.source,
        range: {
          ...validChart.spec.source.range,
          columnCount: 0,
          sheetId: "sheet-1",
        },
      },
      valueDimensions: [],
    },
  };
  const rowLayoutChart: WorkbookChart = {
    ...validChart,
    id: "chart-3",
    spec: {
      ...validChart.spec,
      source: {
        ...validChart.spec.source,
        seriesLayoutBy: "row",
      },
    },
  };

  assert.equal(getWorkbookChartDimensionCount(validChart), 3);
  assert.equal(getWorkbookChartDimensionCount(rowLayoutChart), 4);
  assert.deepEqual(getWorkbookChartValidationIssues(validChart, sheets), []);
  assert.equal(getWorkbookChartStatus(validChart, sheets), "ok");
  assert.deepEqual(createWorkbookChartSummary(validChart, sheets), {
    chartType: "bar",
    id: "chart-1",
    layout: validChart.layout,
    name: "Revenue",
    sheetId: "sheet-1",
    status: "ok",
  });
  assert.deepEqual(
    getWorkbookChartValidationIssues(invalidChart, sheets).map(
      (issue) => issue.code,
    ),
    [
      "CROSS_SHEET_SOURCE",
      "EMPTY_RANGE",
      "MISSING_SHEET",
      "EMPTY_VALUE_DIMENSIONS",
      "INVALID_DIMENSION",
    ],
  );
  assert.equal(getWorkbookChartStatus(invalidChart, sheets), "invalid");
});

test("workbook chart helpers rewrite source ranges for structural row and column edits", () => {
  const baseChart = createBarChart();

  assert.deepEqual(
    adjustWorkbookChartForInsertedRows(baseChart, "sheet-1", 1, 2).spec.source
      .range,
    {
      columnCount: 3,
      rowCount: 4,
      sheetId: "sheet-1",
      startColumn: 1,
      startRow: 4,
    },
  );
  assert.deepEqual(
    adjustWorkbookChartForInsertedRows(baseChart, "sheet-1", 4, 2).spec.source
      .range,
    {
      columnCount: 3,
      rowCount: 6,
      sheetId: "sheet-1",
      startColumn: 1,
      startRow: 2,
    },
  );
  assert.deepEqual(
    adjustWorkbookChartForDeletedRows(baseChart, "sheet-1", 0, 3).spec.source
      .range,
    {
      columnCount: 3,
      rowCount: 3,
      sheetId: "sheet-1",
      startColumn: 1,
      startRow: 0,
    },
  );
  assert.deepEqual(
    adjustWorkbookChartForInsertedColumns(baseChart, "sheet-1", 0, 2).spec
      .source.range,
    {
      columnCount: 3,
      rowCount: 4,
      sheetId: "sheet-1",
      startColumn: 3,
      startRow: 2,
    },
  );
  assert.deepEqual(
    adjustWorkbookChartForDeletedColumns(baseChart, "sheet-1", 2, 5).spec.source
      .range,
    {
      columnCount: 1,
      rowCount: 4,
      sheetId: "sheet-1",
      startColumn: 1,
      startRow: 2,
    },
  );
});

test("workbook chart helpers preserve invalid charts explicitly when structural deletes remove all source data", () => {
  const sheets: WorkbookChartSheetReference[] = [
    {
      columnCount: 10,
      id: "sheet-1",
      rowCount: 12,
    },
  ];
  const chart = createBarChart();
  const withoutRows = adjustWorkbookChartForDeletedRows(
    chart,
    "sheet-1",
    0,
    20,
  );
  const withoutColumns = adjustWorkbookChartForDeletedColumns(
    chart,
    "sheet-1",
    0,
    20,
  );

  assert.deepEqual(withoutRows.spec.source.range, {
    columnCount: 3,
    rowCount: 0,
    sheetId: "sheet-1",
    startColumn: 1,
    startRow: 0,
  });
  assert.deepEqual(withoutColumns.spec.source.range, {
    columnCount: 0,
    rowCount: 4,
    sheetId: "sheet-1",
    startColumn: 0,
    startRow: 2,
  });
  assert.equal(getWorkbookChartStatus(withoutRows, sheets), "invalid");
  assert.equal(getWorkbookChartStatus(withoutColumns, sheets), "invalid");
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
  assert.equal(summary.hasUnsavedChanges, false);
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

test("applyWorkbookTransaction stores sparse cell styles and exposes style ranges", () => {
  const initialState = createWorkbookState();
  const styledState = applyWorkbookTransaction(initialState, {
    operations: [
      {
        columnIndex: 0,
        rowIndex: 0,
        style: {
          bold: true,
          fontSize: 14.9,
          horizontalAlign: "center",
          textColor: "#111827",
        },
        type: "setCellStyle",
      },
      {
        columnCount: 2,
        rowCount: 1,
        startColumn: 1,
        startRow: 0,
        style: {
          backgroundColor: "#f8fafc",
          italic: true,
          wrapText: true,
        },
        type: "setRangeStyle",
      },
    ],
  }).state;

  assert.deepEqual(
    getSheetStyleRange(styledState, {
      columnCount: 3,
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    }).styles,
    [
      [
        {
          bold: true,
          fontSize: 14,
          horizontalAlign: "center",
          textColor: "#111827",
        },
        {
          backgroundColor: "#f8fafc",
          italic: true,
          wrapText: true,
        },
        {
          backgroundColor: "#f8fafc",
          italic: true,
          wrapText: true,
        },
      ],
    ],
  );

  const clearedState = applyWorkbookTransaction(styledState, {
    operations: [
      {
        columnCount: 1,
        rowCount: 1,
        startColumn: 1,
        startRow: 0,
        type: "clearRangeStyle",
      },
    ],
  }).state;

  assert.deepEqual(
    getSheetStyleRange(clearedState, {
      columnCount: 3,
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    }).styles,
    [
      [
        {
          bold: true,
          fontSize: 14,
          horizontalAlign: "center",
          textColor: "#111827",
        },
        null,
        {
          backgroundColor: "#f8fafc",
          italic: true,
          wrapText: true,
        },
      ],
    ],
  );
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
      {
        columnIndex: 1,
        rowIndex: 1,
        style: {
          bold: true,
        },
        type: "setCellStyle",
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
  assert.deepEqual(
    getSheetStyleRange(editedState, {
      columnCount: 3,
      rowCount: 3,
      startColumn: 0,
      startRow: 0,
    }).styles,
    [
      [null, null, null],
      [null, null, null],
      [null, null, { bold: true }],
    ],
  );

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
  assert.deepEqual(getActiveSheet(collapsedState).cellStyles, {});
});

test("applyWorkbookTransaction rewrites persisted chart ranges for structural edits", () => {
  const workbook = createWorkbookState();
  const chart = createBarChart({ sheetId: workbook.activeSheetId });

  chart.spec.source.range.sheetId = workbook.activeSheetId;
  workbook.charts = [chart];
  workbook.nextChartNumber = 2;

  const insertedRows = applyWorkbookTransaction(workbook, {
    operations: [
      {
        count: 2,
        rowIndex: 1,
        sheetId: workbook.activeSheetId,
        type: "insertRows",
      },
    ],
  }).state;
  const deletedColumns = applyWorkbookTransaction(insertedRows, {
    operations: [
      {
        columnIndex: 2,
        count: 1,
        sheetId: workbook.activeSheetId,
        type: "deleteColumns",
      },
    ],
  }).state;

  assert.deepEqual(insertedRows.charts[0]?.spec.source.range, {
    columnCount: 3,
    rowCount: 4,
    sheetId: workbook.activeSheetId,
    startColumn: 1,
    startRow: 4,
  });
  assert.deepEqual(insertedRows.charts[0]?.layout, {
    height: 260,
    offsetX: 0,
    offsetY: 0,
    startColumn: 5,
    startRow: 4,
    width: 420,
    zIndex: 0,
  });
  assert.deepEqual(deletedColumns.charts[0]?.spec.source.range, {
    columnCount: 2,
    rowCount: 4,
    sheetId: workbook.activeSheetId,
    startColumn: 1,
    startRow: 4,
  });
  assert.deepEqual(deletedColumns.charts[0]?.layout, {
    height: 260,
    offsetX: 0,
    offsetY: 0,
    startColumn: 4,
    startRow: 4,
    width: 420,
    zIndex: 0,
  });
});

test("applyWorkbookTransaction preserves charts explicitly when deleting their sheet", () => {
  let workbook = createWorkbookState();
  const primarySheetId = workbook.activeSheetId;

  workbook = applyWorkbookTransaction(workbook, {
    operations: [
      {
        activate: false,
        name: "Sheet 2",
        sheetId: "sheet-2",
        type: "addSheet",
      },
    ],
  }).state;
  const chart = createBarChart({ sheetId: primarySheetId });

  chart.spec.source.range.sheetId = primarySheetId;
  workbook.charts = [chart];
  workbook.nextChartNumber = 2;

  const deletedSheetWorkbook = applyWorkbookTransaction(workbook, {
    operations: [
      {
        nextActiveSheetId: "sheet-2",
        sheetId: primarySheetId,
        type: "deleteSheet",
      },
    ],
  }).state;

  assert.equal(deletedSheetWorkbook.charts.length, 1);
  assert.equal(
    getWorkbookChartStatus(
      deletedSheetWorkbook.charts[0],
      getChartSheetReferences(deletedSheetWorkbook),
    ),
    "invalid",
  );
  assert.deepEqual(
    getWorkbookChartValidationIssues(
      deletedSheetWorkbook.charts[0],
      getChartSheetReferences(deletedSheetWorkbook),
    ).map((issue) => issue.code),
    ["MISSING_SHEET"],
  );
});

test("applyWorkbookTransaction clamps chart layout anchors when sheets shrink", () => {
  const workbook = createWorkbookState();
  const chart = createBarChart({
    layout: {
      height: 260,
      offsetX: 0,
      offsetY: 0,
      startColumn: 20,
      startRow: 80,
      width: 420,
      zIndex: 0,
    },
    sheetId: workbook.activeSheetId,
  });

  chart.spec.source.range.sheetId = workbook.activeSheetId;
  workbook.charts = [chart];
  workbook.nextChartNumber = 2;

  const resizedWorkbook = applyWorkbookTransaction(workbook, {
    operations: [
      {
        columnCount: 3,
        rowCount: 4,
        sheetId: workbook.activeSheetId,
        type: "resizeSheet",
      },
    ],
  }).state;

  assert.deepEqual(resizedWorkbook.charts[0]?.layout, {
    height: 260,
    offsetX: 0,
    offsetY: 0,
    startColumn: 2,
    startRow: 3,
    width: 420,
    zIndex: 0,
  });
});

test("applyWorkbookTransaction manages chart lifecycle operations", () => {
  const initialState = applyWorkbookTransaction(createWorkbookState(), {
    operations: [
      {
        activate: false,
        columnCount: 4,
        name: "Metrics",
        rowCount: 6,
        sheetId: "sheet-metrics",
        type: "addSheet",
      },
    ],
  }).state;
  const primarySheetId = initialState.activeSheetId;

  const afterAdd = applyWorkbookTransaction(initialState, {
    operations: [
      {
        spec: {
          categoryDimension: 0,
          chartType: "bar",
          family: "cartesian",
          source: {
            range: {
              columnCount: 2,
              rowCount: 4,
              sheetId: primarySheetId,
              startColumn: 0,
              startRow: 0,
            },
            seriesLayoutBy: "column",
            sourceHeader: true,
          },
          valueDimensions: [1],
        },
        type: "addChart",
      },
    ],
  }).state;

  assert.equal(afterAdd.nextChartNumber, 2);
  assert.deepEqual(afterAdd.charts, [
    {
      id: "chart-1",
      layout: {
        height: 260,
        offsetX: 0,
        offsetY: 0,
        startColumn: 3,
        startRow: 0,
        width: 420,
        zIndex: 0,
      },
      name: "Chart 1",
      sheetId: primarySheetId,
      spec: {
        categoryDimension: 0,
        chartType: "bar",
        family: "cartesian",
        source: {
          range: {
            columnCount: 2,
            rowCount: 4,
            sheetId: primarySheetId,
            startColumn: 0,
            startRow: 0,
          },
          seriesLayoutBy: "column",
          sourceHeader: true,
        },
        valueDimensions: [1],
      },
    },
  ]);
  assert.deepEqual(getWorkbookSummary(afterAdd).charts, [
    {
      chartType: "bar",
      id: "chart-1",
      layout: {
        height: 260,
        offsetX: 0,
        offsetY: 0,
        startColumn: 3,
        startRow: 0,
        width: 420,
        zIndex: 0,
      },
      name: "Chart 1",
      sheetId: primarySheetId,
      status: "ok",
    },
  ]);

  const afterLayout = applyWorkbookTransaction(afterAdd, {
    operations: [
      {
        chartId: "chart-1",
        layout: {
          height: 300,
          offsetX: 12,
          offsetY: 8,
          startColumn: 2,
          startRow: 1,
          width: 500,
          zIndex: 4,
        },
        type: "setChartLayout",
      },
    ],
  }).state;

  assert.deepEqual(afterLayout.charts[0]?.layout, {
    height: 300,
    offsetX: 12,
    offsetY: 8,
    startColumn: 2,
    startRow: 1,
    width: 500,
    zIndex: 4,
  });

  const afterRenameAndRetarget = applyWorkbookTransaction(afterLayout, {
    operations: [
      {
        chartId: "chart-1",
        name: "  Margin Mix  ",
        type: "renameChart",
      },
      {
        chartId: "chart-1",
        spec: {
          chartType: "pie",
          family: "pie",
          nameDimension: 0,
          source: {
            range: {
              columnCount: 2,
              rowCount: 4,
              sheetId: "sheet-metrics",
              startColumn: 1,
              startRow: 2,
            },
            seriesLayoutBy: "column",
            sourceHeader: true,
          },
          valueDimension: 1,
        },
        type: "setChartSpec",
      },
    ],
  }).state;

  assert.equal(afterRenameAndRetarget.charts[0]?.name, "Margin Mix");
  assert.equal(afterRenameAndRetarget.charts[0]?.sheetId, "sheet-metrics");
  assert.equal(afterRenameAndRetarget.charts[0]?.spec.family, "pie");
  assert.deepEqual(afterRenameAndRetarget.charts[0]?.layout, {
    height: 300,
    offsetX: 12,
    offsetY: 8,
    startColumn: 2,
    startRow: 1,
    width: 500,
    zIndex: 4,
  });
  assert.equal(afterRenameAndRetarget.nextChartNumber, 2);

  const afterDelete = applyWorkbookTransaction(afterRenameAndRetarget, {
    operations: [
      {
        chartId: "chart-1",
        type: "deleteChart",
      },
    ],
  }).state;

  assert.equal(afterDelete.charts.length, 0);
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
  assert.throws(
    () =>
      applyWorkbookTransaction(initialState, {
        operations: [
          {
            chartId: "chart-1",
            spec: {
              categoryDimension: 0,
              chartType: "bar",
              family: "cartesian",
              source: {
                range: {
                  columnCount: 0,
                  rowCount: 1,
                  sheetId: initialState.activeSheetId,
                  startColumn: 0,
                  startRow: 0,
                },
                seriesLayoutBy: "column",
                sourceHeader: true,
              },
              valueDimensions: [1],
            },
            type: "addChart",
          },
        ],
      }),
    /Chart "chart-1" cannot be added:/,
  );

  const stateWithChart = applyWorkbookTransaction(initialState, {
    operations: [
      {
        chartId: "chart-1",
        spec: {
          categoryDimension: 0,
          chartType: "bar",
          family: "cartesian",
          source: {
            range: {
              columnCount: 2,
              rowCount: 4,
              sheetId: initialState.activeSheetId,
              startColumn: 0,
              startRow: 0,
            },
            seriesLayoutBy: "column",
            sourceHeader: true,
          },
          valueDimensions: [1],
        },
        type: "addChart",
      },
    ],
  }).state;

  assert.throws(
    () =>
      applyWorkbookTransaction(stateWithChart, {
        operations: [
          {
            chartId: "chart-1",
            type: "deleteChart",
          },
          {
            chartId: "chart-1",
            type: "deleteChart",
          },
        ],
      }),
    /Chart "chart-1" was not found\./,
  );
  assert.throws(
    () =>
      applyWorkbookTransaction(stateWithChart, {
        operations: [
          {
            chartId: "chart-1",
            spec: {
              chartType: "pie",
              family: "pie",
              nameDimension: 0,
              source: {
                range: {
                  columnCount: 1,
                  rowCount: 4,
                  sheetId: initialState.activeSheetId,
                  startColumn: 0,
                  startRow: 0,
                },
                seriesLayoutBy: "column",
                sourceHeader: true,
              },
              valueDimension: 1,
            },
            type: "setChartSpec",
          },
        ],
      }),
    /Chart "chart-1" cannot be updated:/,
  );
});
