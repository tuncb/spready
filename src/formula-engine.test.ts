import assert from "node:assert/strict";
import { test } from "node:test";

import {
  evaluateSheet,
  getCellEvaluation,
  type CellKey,
} from "./formula-engine";
import { normalizeSheet, type WorkbookSheet } from "./workbook-core";

function createSheet(rows: string[][]): WorkbookSheet {
  return {
    cells: normalizeSheet(rows),
    id: "sheet-under-test",
    name: "Sheet Under Test",
  };
}

function getDisplay(
  sheet: WorkbookSheet,
  rowIndex: number,
  columnIndex: number,
) {
  return getCellEvaluation(evaluateSheet(sheet, 1), rowIndex, columnIndex)
    .display;
}

function getDependents(snapshotCellKeys: Iterable<CellKey>) {
  return [...snapshotCellKeys].sort();
}

test("evaluateSheet handles arithmetic precedence, parentheses, unary operators, and references", () => {
  const sheet = createSheet([
    ["1", "2", "=1+2*3", "=(1+2)*3", "=A1+B1", "=-(A1+B1)"],
    ["", "", "=A2+5", "=a1+b1", "", ""],
  ]);
  const snapshot = evaluateSheet(sheet, 7);

  assert.equal(getCellEvaluation(snapshot, 0, 2).display, "7");
  assert.equal(getCellEvaluation(snapshot, 0, 3).display, "9");
  assert.equal(getCellEvaluation(snapshot, 0, 4).display, "3");
  assert.equal(getCellEvaluation(snapshot, 0, 5).display, "-3");
  assert.equal(getCellEvaluation(snapshot, 1, 2).display, "5");
  assert.equal(getCellEvaluation(snapshot, 1, 3).display, "3");
});

test("evaluateSheet returns formula error markers for parse, reference, divide by zero, value, and cycle failures", () => {
  const sheet = createSheet([
    ["text", "=1+", "=Z99", "=1/0", "=A1+1", "=G1", "=F1"],
  ]);

  assert.equal(getDisplay(sheet, 0, 1), "#ERROR!");
  assert.equal(getDisplay(sheet, 0, 2), "#REF!");
  assert.equal(getDisplay(sheet, 0, 3), "#DIV/0!");
  assert.equal(getDisplay(sheet, 0, 4), "#VALUE!");
  assert.equal(getDisplay(sheet, 0, 5), "#CYCLE!");
  assert.equal(getDisplay(sheet, 0, 6), "#CYCLE!");
});

test("evaluateSheet records direct precedents and dependents for formula cells", () => {
  const sheet = createSheet([["1", "2", "=A1+B1", "=C1*2"]]);
  const snapshot = evaluateSheet(sheet, 3);

  assert.deepEqual(getCellEvaluation(snapshot, 0, 2).dependencies, [
    "0:0",
    "0:1",
  ]);
  assert.deepEqual(getCellEvaluation(snapshot, 0, 3).dependencies, ["0:2"]);
  assert.deepEqual(getDependents(snapshot.precedents.get("0:2") ?? []), [
    "0:0",
    "0:1",
  ]);
  assert.deepEqual(getDependents(snapshot.dependents.get("0:2") ?? []), [
    "0:3",
  ]);
});
