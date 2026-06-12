import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createSortTableOperation,
  getVisibleTableHeaderSortTargets,
} from "./app-table-sort-controls";
import type { WorkbookTableSummary } from "./workbook-core";

const baseTable: WorkbookTableSummary = {
  hasHeaderRow: true,
  id: "table-sales",
  name: "Sales",
  range: {
    columnCount: 3,
    rowCount: 4,
    sheetId: "sheet-1",
    startColumn: 1,
    startRow: 2,
  },
  sheetId: "sheet-1",
};

test("getVisibleTableHeaderSortTargets returns visible table header columns", () => {
  assert.deepEqual(
    getVisibleTableHeaderSortTargets(
      [
        {
          ...baseTable,
          sortState: {
            keys: [
              {
                columnIndex: 2,
                direction: "descending",
              },
            ],
            valueMode: "display",
          },
        },
      ],
      {
        height: 3,
        width: 2,
        x: 2,
        y: 1,
      },
    ),
    [
      {
        columnIndex: 2,
        direction: "descending",
        rowIndex: 2,
        tableId: "table-sales",
      },
      {
        columnIndex: 3,
        direction: undefined,
        rowIndex: 2,
        tableId: "table-sales",
      },
    ],
  );
});

test("getVisibleTableHeaderSortTargets ignores hidden and non-header table rows", () => {
  assert.deepEqual(
    getVisibleTableHeaderSortTargets(
      [
        baseTable,
        {
          ...baseTable,
          hasHeaderRow: false,
          id: "table-no-header",
        },
      ],
      {
        height: 3,
        width: 4,
        x: 0,
        y: 3,
      },
    ),
    [],
  );
});

test("createSortTableOperation builds the shared table sort transaction", () => {
  assert.deepEqual(createSortTableOperation("table-sales", 2, "ascending"), {
    keys: [
      {
        columnIndex: 2,
        direction: "ascending",
      },
    ],
    tableId: "table-sales",
    type: "sortTable",
    valueMode: "display",
  });
});
