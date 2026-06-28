import assert from "node:assert/strict";
import { test } from "node:test";

import { getTableColumnHighlightStyle } from "./app-table-column-highlights";
import type { WorkbookTableSummary } from "./workbook-core";

const table: WorkbookTableSummary = {
  columnHighlightRules: [
    {
      columnIndex: 2,
      threshold: 100,
    },
  ],
  hasHeaderRow: true,
  id: "table-sales",
  name: "Sales",
  range: {
    columnCount: 3,
    rowCount: 4,
    sheetId: "sheet-1",
    startColumn: 0,
    startRow: 1,
  },
  sheetId: "sheet-1",
};

test("getTableColumnHighlightStyle highlights table body values above the threshold", () => {
  assert.deepEqual(getTableColumnHighlightStyle([table], 2, 2, "125"), {
    backgroundColor: "#fff1c2",
    bold: true,
    textColor: "#7c2d12",
  });
  assert.equal(getTableColumnHighlightStyle([table], 2, 2, "100"), undefined);
  assert.equal(getTableColumnHighlightStyle([table], 2, 1, "125"), undefined);
  assert.equal(getTableColumnHighlightStyle([table], 1, 2, "125"), undefined);
});

test("getTableColumnHighlightStyle supports threshold comparison types", () => {
  const createTable = (
    thresholdType: NonNullable<
      WorkbookTableSummary["columnHighlightRules"]
    >[number]["thresholdType"],
  ): WorkbookTableSummary => ({
    ...table,
    columnHighlightRules: [
      {
        columnIndex: 2,
        threshold: 100,
        thresholdType,
      },
    ],
  });

  assert.equal(getTableColumnHighlightStyle([createTable("lower")], 2, 2, "99")?.bold, true);
  assert.equal(getTableColumnHighlightStyle([createTable("lower")], 2, 2, "100"), undefined);
  assert.equal(
    getTableColumnHighlightStyle([createTable("lowerOrEqual")], 2, 2, "100")?.bold,
    true,
  );
  assert.equal(getTableColumnHighlightStyle([createTable("equal")], 2, 2, "100")?.bold, true);
  assert.equal(getTableColumnHighlightStyle([createTable("equal")], 2, 2, "101"), undefined);
  assert.equal(
    getTableColumnHighlightStyle([createTable("higherOrEqual")], 2, 2, "100")?.bold,
    true,
  );
  assert.equal(getTableColumnHighlightStyle([createTable("higher")], 2, 2, "100"), undefined);
});

test("getTableColumnHighlightStyle merges matching rules on one column in order", () => {
  assert.deepEqual(
    getTableColumnHighlightStyle(
      [
        {
          ...table,
          columnHighlightRules: [
            {
              backgroundColor: "#fee2e2",
              columnIndex: 2,
              threshold: 100,
              thresholdType: "higherOrEqual",
            },
            {
              columnIndex: 2,
              textColor: "#991b1b",
              threshold: 150,
              thresholdType: "higherOrEqual",
            },
          ],
        },
      ],
      2,
      2,
      "175",
    ),
    {
      backgroundColor: "#fee2e2",
      bold: true,
      textColor: "#991b1b",
    },
  );
});

test("getTableColumnHighlightStyle uses configured highlight styling", () => {
  assert.deepEqual(
    getTableColumnHighlightStyle(
      [
        {
          ...table,
          columnHighlightRules: [
            {
              backgroundColor: "#fde68a",
              bold: false,
              columnIndex: 2,
              textColor: "#111827",
              threshold: 1000,
            },
          ],
        },
      ],
      2,
      2,
      "1,200",
    ),
    {
      backgroundColor: "#fde68a",
      bold: false,
      textColor: "#111827",
    },
  );
});
