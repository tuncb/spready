import assert from "node:assert/strict";
import path from "node:path";
import { test } from "node:test";

import type { WorkbookSummary } from "./workbook-core";
import { formatWorkbookWindowTitle } from "./window-title";

function createSummary(overrides: Partial<WorkbookSummary> = {}): WorkbookSummary {
  return {
    activeSheetId: "sheet-1",
    activeSheetName: "Sheet 1",
    charts: [],
    hasUnsavedChanges: false,
    sheets: [
      {
        columnCount: 10,
        id: "sheet-1",
        name: "Sheet 1",
        rowCount: 20,
      },
    ],
    version: 0,
    ...overrides,
  };
}

test("formatWorkbookWindowTitle uses the workbook file name when available", () => {
  const title = formatWorkbookWindowTitle(
    createSummary({
      documentFilePath: path.join("workbooks", "budget.spready"),
    }),
    "Spready",
  );

  assert.equal(title, "Spready - budget.spready");
});

test("formatWorkbookWindowTitle prefixes dirty workbooks with an asterisk", () => {
  const title = formatWorkbookWindowTitle(
    createSummary({
      documentFilePath: path.join("workbooks", "budget.spready"),
      hasUnsavedChanges: true,
    }),
    "Spready",
  );

  assert.equal(title, "Spready - *budget.spready");
});

test("formatWorkbookWindowTitle falls back to Untitled for unsaved workbooks", () => {
  const title = formatWorkbookWindowTitle(createSummary(), "Spready");

  assert.equal(title, "Spready - Untitled");
});

test("formatWorkbookWindowTitle shows a dirty untitled workbook clearly", () => {
  const title = formatWorkbookWindowTitle(
    createSummary({
      hasUnsavedChanges: true,
    }),
    "Spready",
  );

  assert.equal(title, "Spready - *Untitled");
});
