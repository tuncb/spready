import assert from "node:assert/strict";
import { test } from "node:test";

import { applyWorkbookTransaction, createWorkbookState } from "./workbook-core";
import {
  parseWorkbookDocument,
  serializeWorkbookDocument,
  WORKBOOK_DOCUMENT_FORMAT,
  WORKBOOK_DOCUMENT_VERSION,
} from "./workbook-document";

test("workbook documents round-trip sparse multi-sheet workbook state", () => {
  const workbook = applyWorkbookTransaction(createWorkbookState(), {
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["Revenue", "2026", "=B2*2"],
          ["North", "1200", ""],
        ],
      },
      {
        sourceFilePath: "C:\\imports\\quarterly.csv",
        type: "setSheetSourceFile",
      },
      {
        columnCount: 2,
        rowCount: 1,
        startColumn: 0,
        startRow: 0,
        style: {
          bold: true,
          fontSize: 16,
          textColor: "#0f172a",
        },
        type: "setRangeStyle",
      },
      {
        activate: true,
        columnCount: 3,
        name: "Budget",
        rowCount: 4,
        sheetId: "sheet-12",
        type: "addSheet",
      },
      {
        columnIndex: 0,
        rowIndex: 0,
        sheetId: "sheet-12",
        type: "setCell",
        value: "2026",
      },
      {
        columnIndex: 1,
        rowIndex: 0,
        sheetId: "sheet-12",
        type: "setCell",
        value: "=A1+1",
      },
      {
        columnIndex: 1,
        rowIndex: 0,
        sheetId: "sheet-12",
        style: {
          italic: true,
          horizontalAlign: "right",
        },
        type: "setCellStyle",
      },
    ],
  }).state;

  workbook.documentFilePath = "C:\\workbooks\\budget.spready";
  workbook.charts = [
    {
      id: "chart-1",
      layout: {
        height: 260,
        offsetX: 0,
        offsetY: 0,
        startColumn: 4,
        startRow: 0,
        width: 420,
        zIndex: 0,
      },
      name: "Quarterly Revenue",
      sheetId: workbook.sheets[0].id,
      spec: {
        categoryDimension: 0,
        chartType: "bar",
        family: "cartesian",
        source: {
          range: {
            columnCount: 2,
            rowCount: 2,
            sheetId: workbook.sheets[0].id,
            startColumn: 0,
            startRow: 0,
          },
          seriesLayoutBy: "column",
          sourceHeader: true,
        },
        valueDimensions: [1],
      },
    },
    {
      id: "chart-2",
      layout: {
        height: 260,
        offsetX: 0,
        offsetY: 0,
        startColumn: 0,
        startRow: 0,
        width: 420,
        zIndex: 1,
      },
      name: "Broken Imported Chart",
      sheetId: "missing-sheet",
      spec: {
        chartType: "pie",
        family: "pie",
        nameDimension: 0,
        source: {
          range: {
            columnCount: 0,
            rowCount: 0,
            sheetId: "missing-sheet",
            startColumn: 0,
            startRow: 0,
          },
          seriesLayoutBy: "column",
          sourceHeader: false,
        },
        valueDimension: 0,
      },
    },
  ];
  workbook.nextChartNumber = 3;

  const serialized = serializeWorkbookDocument(workbook);
  const parsed = parseWorkbookDocument(serialized);

  assert.ok(serialized.includes(`"format": "${WORKBOOK_DOCUMENT_FORMAT}"`));
  assert.ok(serialized.includes(`"formatVersion": ${WORKBOOK_DOCUMENT_VERSION}`));
  assert.ok(!serialized.includes("documentFilePath"));
  assert.equal(parsed.version, 0);
  assert.equal(parsed.documentFilePath, undefined);
  assert.equal(parsed.activeSheetId, "sheet-12");
  assert.equal(parsed.nextChartNumber, 3);
  assert.equal(parsed.nextSheetNumber, workbook.nextSheetNumber);
  assert.deepEqual(parsed.charts, workbook.charts);
  assert.deepEqual(
    parsed.sheets.map((sheet) => ({
      columnCount: sheet.cells[0]?.length ?? 0,
      id: sheet.id,
      name: sheet.name,
      rowCount: sheet.cells.length,
      sourceFilePath: sheet.sourceFilePath,
    })),
    [
      {
        columnCount: 50,
        id: workbook.sheets[0].id,
        name: "Sheet 1",
        rowCount: 200,
        sourceFilePath: "C:\\imports\\quarterly.csv",
      },
      {
        columnCount: 3,
        id: "sheet-12",
        name: "Budget",
        rowCount: 4,
        sourceFilePath: undefined,
      },
    ],
  );
  assert.deepEqual(parsed.sheets[0].cells[0].slice(0, 3), ["Revenue", "2026", "=B2*2"]);
  assert.deepEqual(parsed.sheets[0].cells[1].slice(0, 3), ["North", "1200", ""]);
  assert.deepEqual(parsed.sheets[0].cellStyles, {
    "0:0": {
      bold: true,
      fontSize: 16,
      textColor: "#0f172a",
    },
    "0:1": {
      bold: true,
      fontSize: 16,
      textColor: "#0f172a",
    },
  });
  assert.deepEqual(parsed.sheets[1].cells[0].slice(0, 2), ["2026", "=A1+1"]);
  assert.deepEqual(parsed.sheets[1].cellStyles, {
    "0:1": {
      horizontalAlign: "right",
      italic: true,
    },
  });
  assert.equal(parsed.sheets[1].cells[3][2], "");
});

test("workbook documents reject invalid workbook references and cell entries", () => {
  assert.throws(
    () =>
      parseWorkbookDocument(
        JSON.stringify({
          format: WORKBOOK_DOCUMENT_FORMAT,
          formatVersion: WORKBOOK_DOCUMENT_VERSION,
          workbook: {
            activeSheetId: "missing",
            charts: [],
            nextChartNumber: 1,
            nextSheetNumber: 2,
            sheets: [
              {
                cells: [],
                columnCount: 2,
                id: "sheet-1",
                name: "Sheet 1",
                rowCount: 2,
                styles: [],
              },
            ],
          },
        }),
      ),
    /missing active sheet "missing"/,
  );

  assert.throws(
    () =>
      parseWorkbookDocument(
        JSON.stringify({
          format: WORKBOOK_DOCUMENT_FORMAT,
          formatVersion: WORKBOOK_DOCUMENT_VERSION,
          workbook: {
            activeSheetId: "sheet-1",
            charts: [],
            nextChartNumber: 1,
            nextSheetNumber: 2,
            sheets: [
              {
                cells: [{ column: 4, row: 0, value: "bad" }],
                columnCount: 2,
                id: "sheet-1",
                name: "Sheet 1",
                rowCount: 2,
                styles: [],
              },
            ],
          },
        }),
      ),
    /out-of-bounds cell 0:4/,
  );

  assert.throws(
    () =>
      parseWorkbookDocument(
        JSON.stringify({
          format: WORKBOOK_DOCUMENT_FORMAT,
          formatVersion: WORKBOOK_DOCUMENT_VERSION,
          workbook: {
            activeSheetId: "sheet-1",
            charts: [
              {
                id: "chart-1",
                layout: {
                  height: 260,
                  offsetX: 0,
                  offsetY: 0,
                  startColumn: 0,
                  startRow: 0,
                  width: 420,
                  zIndex: 0,
                },
                name: "Chart A",
                sheetId: "sheet-1",
                spec: {
                  categoryDimension: 0,
                  chartType: "bar",
                  family: "cartesian",
                  source: {
                    range: {
                      columnCount: 2,
                      rowCount: 2,
                      sheetId: "sheet-1",
                      startColumn: 0,
                      startRow: 0,
                    },
                    seriesLayoutBy: "column",
                    sourceHeader: true,
                  },
                  valueDimensions: [1],
                },
              },
              {
                id: "chart-1",
                layout: {
                  height: 260,
                  offsetX: 0,
                  offsetY: 0,
                  startColumn: 0,
                  startRow: 0,
                  width: 420,
                  zIndex: 1,
                },
                name: "Chart B",
                sheetId: "sheet-1",
                spec: {
                  chartType: "pie",
                  family: "pie",
                  nameDimension: 0,
                  source: {
                    range: {
                      columnCount: 2,
                      rowCount: 2,
                      sheetId: "sheet-1",
                      startColumn: 0,
                      startRow: 0,
                    },
                    seriesLayoutBy: "column",
                    sourceHeader: true,
                  },
                  valueDimension: 1,
                },
              },
            ],
            nextChartNumber: 2,
            nextSheetNumber: 2,
            sheets: [
              {
                cells: [],
                columnCount: 2,
                id: "sheet-1",
                name: "Sheet 1",
                rowCount: 2,
                styles: [],
              },
            ],
          },
        }),
      ),
    /duplicate chart id "chart-1"/,
  );
});
