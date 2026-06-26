import assert from "node:assert/strict";
import { test } from "node:test";

import { filterSheetQuickOpenResults } from "./sheet-quick-open";
import type { SheetSummary } from "./workbook-core";

const sheets: SheetSummary[] = [
  createSheet("summary", "Summary"),
  createSheet("north-sales", "North Sales"),
  createSheet("south-sales", "South Sales"),
  createSheet("sales-archive", "Sales Archive"),
  createSheet("forecast", "2026 Forecast"),
];

test("filterSheetQuickOpenResults returns all sheets for an empty query", () => {
  assert.deepEqual(
    filterSheetQuickOpenResults(sheets, "").map((sheet) => sheet.id),
    ["summary", "north-sales", "south-sales", "sales-archive", "forecast"],
  );
});

test("filterSheetQuickOpenResults ranks exact, prefix, word-prefix, then substring matches", () => {
  assert.deepEqual(
    filterSheetQuickOpenResults(sheets, "sales").map((sheet) => sheet.id),
    ["sales-archive", "north-sales", "south-sales"],
  );
});

test("filterSheetQuickOpenResults trims and ignores case", () => {
  assert.deepEqual(
    filterSheetQuickOpenResults(sheets, "  SOUTH  ").map((sheet) => sheet.id),
    ["south-sales"],
  );
});

test("filterSheetQuickOpenResults returns no results when no sheet matches", () => {
  assert.deepEqual(filterSheetQuickOpenResults(sheets, "budget"), []);
});

function createSheet(id: string, name: string): SheetSummary {
  return {
    columnCount: 5,
    id,
    name,
    rowCount: 10,
  };
}
