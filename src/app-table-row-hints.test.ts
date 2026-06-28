import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildTableRowHintItems,
  buildTableRowHintRangeRequests,
  getTableRowHintTarget,
  getTableRowHintTableKey,
  getTableRowHintTargetKey,
} from "./app-table-row-hints";
import type { WorkbookTableSummary } from "./workbook-core";

const baseTable: WorkbookTableSummary = {
  hasHeaderRow: true,
  id: "table-sales",
  name: "Sales",
  range: {
    columnCount: 3,
    rowCount: 5,
    sheetId: "sheet-1",
    startColumn: 1,
    startRow: 2,
  },
  sheetId: "sheet-1",
};

test("getTableRowHintTarget returns body cells inside headered tables", () => {
  assert.deepEqual(getTableRowHintTarget([baseTable], [2, 4]), {
    activeColumnIndex: 2,
    rowIndex: 4,
    table: baseTable,
    tableRowNumber: 2,
  });
});

test("getTableRowHintTarget ignores header cells and cells outside the table", () => {
  assert.equal(getTableRowHintTarget([baseTable], [2, 2]), null);
  assert.equal(getTableRowHintTarget([baseTable], [4, 4]), null);
  assert.equal(getTableRowHintTarget([baseTable], [2, 7]), null);
});

test("getTableRowHintTarget ignores tables without header rows", () => {
  assert.equal(
    getTableRowHintTarget(
      [
        {
          ...baseTable,
          hasHeaderRow: false,
        },
      ],
      [2, 4],
    ),
    null,
  );
});

test("buildTableRowHintRangeRequests requests the table header and selected row", () => {
  const target = getTableRowHintTarget([baseTable], [2, 4]);

  assert.ok(target);
  assert.deepEqual(buildTableRowHintRangeRequests(target), {
    header: {
      columnCount: 3,
      rowCount: 1,
      sheetId: "sheet-1",
      startColumn: 1,
      startRow: 2,
    },
    row: {
      columnCount: 3,
      rowCount: 1,
      sheetId: "sheet-1",
      startColumn: 1,
      startRow: 4,
    },
  });
});

test("buildTableRowHintItems pairs column names with row values and marks the active cell", () => {
  const target = getTableRowHintTarget([baseTable], [2, 4]);

  assert.ok(target);
  assert.deepEqual(buildTableRowHintItems(target, ["Region", "", "Total"], ["West", "Q2", "12"]), [
    {
      columnIndex: 1,
      isActive: false,
      label: "Region",
      value: "West",
    },
    {
      columnIndex: 2,
      isActive: true,
      label: "C",
      value: "Q2",
    },
    {
      columnIndex: 3,
      isActive: false,
      label: "Total",
      value: "12",
    },
  ]);
});

test("getTableRowHintTargetKey scopes row hints by target and workbook version", () => {
  const target = getTableRowHintTarget([baseTable], [2, 4]);

  assert.ok(target);
  assert.equal(getTableRowHintTargetKey(target, 9), "sheet-1:table-sales:4:2:9");
});

test("getTableRowHintTableKey scopes scroll stability by table", () => {
  const target = getTableRowHintTarget([baseTable], [2, 4]);

  assert.ok(target);
  assert.equal(getTableRowHintTableKey(target), "sheet-1:table-sales");
});
