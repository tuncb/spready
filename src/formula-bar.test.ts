import assert from "node:assert/strict";
import test from "node:test";

import { getFormulaBarPreview } from "./formula-bar";

test("returns idle when no cell data is available", () => {
  assert.equal(getFormulaBarPreview(null), "idle");
});

test("returns the literal cell display instead of a type label", () => {
  assert.equal(
    getFormulaBarPreview({
      columnIndex: 1,
      display: "hello",
      input: "hello",
      isFormula: false,
      rowIndex: 2,
      sheetId: "sheet-1",
      sheetName: "Sheet 1",
    }),
    "hello",
  );
});

test("returns the evaluated display for formula cells", () => {
  assert.equal(
    getFormulaBarPreview({
      columnIndex: 1,
      display: "42",
      input: "=A1+B1",
      isFormula: true,
      rowIndex: 2,
      sheetId: "sheet-1",
      sheetName: "Sheet 1",
    }),
    "42",
  );
});
