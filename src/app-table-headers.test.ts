import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildStickyTableHeaderRangeRequest,
  getStickyTableHeader,
  getTableHeaderCacheKey,
} from "./app-table-headers";
import type { WorkbookTableSummary } from "./workbook-core";

const baseTable: WorkbookTableSummary = {
  hasHeaderRow: true,
  id: "table-sales",
  name: "Sales",
  range: {
    columnCount: 4,
    rowCount: 10,
    sheetId: "sheet-1",
    startColumn: 2,
    startRow: 5,
  },
  sheetId: "sheet-1",
};

test("getStickyTableHeader returns the visible table after its header scrolls off", () => {
  assert.equal(
    getStickyTableHeader([baseTable], {
      height: 20,
      width: 3,
      x: 3,
      y: 8,
    }),
    baseTable,
  );
});

test("getStickyTableHeader ignores visible header rows and rows outside the table", () => {
  assert.equal(
    getStickyTableHeader([baseTable], {
      height: 20,
      width: 3,
      x: 3,
      y: 5,
    }),
    null,
  );

  assert.equal(
    getStickyTableHeader([baseTable], {
      height: 20,
      width: 3,
      x: 3,
      y: 15,
    }),
    null,
  );
});

test("getStickyTableHeader ignores tables without visible table columns", () => {
  assert.equal(
    getStickyTableHeader([baseTable], {
      height: 20,
      width: 2,
      x: 6,
      y: 8,
    }),
    null,
  );
});

test("buildStickyTableHeaderRangeRequest returns the visible header column span", () => {
  assert.deepEqual(
    buildStickyTableHeaderRangeRequest(baseTable, {
      height: 20,
      width: 5,
      x: 0,
      y: 8,
    }),
    {
      columnCount: 3,
      rowCount: 1,
      sheetId: "sheet-1",
      startColumn: 2,
      startRow: 5,
    },
  );
});

test("getTableHeaderCacheKey scopes cached header cells by sheet, table, and version", () => {
  assert.equal(getTableHeaderCacheKey("sheet-1", "table-sales", 12), "sheet-1:table-sales:12");
});
