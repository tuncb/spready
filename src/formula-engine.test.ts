import assert from "node:assert/strict";
import { test } from "node:test";

import { evaluateSheet, evaluateWorkbook, getCellEvaluation, type CellKey } from "./formula-engine";
import { normalizeSheet, type WorkbookSheet } from "./workbook-core";

function createSheet(rows: string[][]): WorkbookSheet {
  return {
    cells: normalizeSheet(rows),
    cellStyles: {},
    id: "sheet-under-test",
    name: "Sheet Under Test",
  };
}

function getDisplay(sheet: WorkbookSheet, rowIndex: number, columnIndex: number) {
  return getCellEvaluation(evaluateSheet(sheet, 1), rowIndex, columnIndex).display;
}

function getDependents(snapshotCellKeys: Iterable<CellKey>) {
  return [...snapshotCellKeys].sort();
}

function cellKey(rowIndex: number, columnIndex: number, sheetId = "sheet-under-test") {
  return `${sheetId}:${rowIndex}:${columnIndex}`;
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

test("evaluateSheet returns formula error markers for parse, name, reference, divide by zero, value, and cycle failures", () => {
  const sheet = createSheet([
    ["text", "=1+", "=MissingName", "=Z99", "=1/0", "=A1+1", "=H1", "=G1"],
  ]);

  assert.equal(getDisplay(sheet, 0, 1), "#ERROR!");
  assert.equal(getDisplay(sheet, 0, 2), "#NAME?");
  assert.equal(getDisplay(sheet, 0, 3), "#REF!");
  assert.equal(getDisplay(sheet, 0, 4), "#DIV/0!");
  assert.equal(getDisplay(sheet, 0, 5), "#VALUE!");
  assert.equal(getDisplay(sheet, 0, 6), "#CYCLE!");
  assert.equal(getDisplay(sheet, 0, 7), "#CYCLE!");
});

test("evaluateSheet records direct precedents and dependents for formula cells", () => {
  const sheet = createSheet([["1", "2", "=A1+B1", "=C1*2"]]);
  const snapshot = evaluateSheet(sheet, 3);

  assert.deepEqual(getCellEvaluation(snapshot, 0, 2).dependencies, [cellKey(0, 0), cellKey(0, 1)]);
  assert.deepEqual(getCellEvaluation(snapshot, 0, 3).dependencies, [cellKey(0, 2)]);
  assert.deepEqual(getDependents(snapshot.precedents.get(cellKey(0, 2)) ?? []), [
    cellKey(0, 0),
    cellKey(0, 1),
  ]);
  assert.deepEqual(getDependents(snapshot.dependents.get(cellKey(0, 2)) ?? []), [cellKey(0, 3)]);
});

test("evaluateSheet supports text, boolean, comparison, exponent, percent, and error literals", () => {
  const sheet = createSheet([
    ['= "Hello" & " " & "World"', "=TRUE", "=FALSE", "=2^3", "=50%", '="A"="a"', "=1<>2", "=#N/A"],
  ]);
  const snapshot = evaluateSheet(sheet, 11);

  assert.equal(getCellEvaluation(snapshot, 0, 0).display, "Hello World");
  assert.equal(getCellEvaluation(snapshot, 0, 1).display, "TRUE");
  assert.equal(getCellEvaluation(snapshot, 0, 2).display, "FALSE");
  assert.equal(getCellEvaluation(snapshot, 0, 3).display, "8");
  assert.equal(getCellEvaluation(snapshot, 0, 4).display, "0.5");
  assert.equal(getCellEvaluation(snapshot, 0, 5).display, "TRUE");
  assert.equal(getCellEvaluation(snapshot, 0, 6).display, "TRUE");
  assert.equal(getCellEvaluation(snapshot, 0, 7).display, "#N/A");
});

test("evaluateSheet treats single-cell ranges as scalars and multi-cell ranges as value errors outside functions", () => {
  const sheet = createSheet([["1", "2", "=A1:A1", "=A1:B1", "=A1:A1+4", "=A1:B1+1"]]);
  const snapshot = evaluateSheet(sheet, 19);

  assert.equal(getCellEvaluation(snapshot, 0, 2).display, "1");
  assert.equal(getCellEvaluation(snapshot, 0, 3).display, "#VALUE!");
  assert.equal(getCellEvaluation(snapshot, 0, 4).display, "5");
  assert.equal(getCellEvaluation(snapshot, 0, 5).display, "#VALUE!");
  assert.deepEqual(getCellEvaluation(snapshot, 0, 3).dependencies, [cellKey(0, 0), cellKey(0, 1)]);
});

test("evaluateSheet supports Excel-style reference intersections", () => {
  const sheet = createSheet([
    [
      "1",
      "2",
      "3",
      "=SUM(A1:C1 B1:D1)",
      "=B1:C1 A1:D1",
      "=A1:A2 C1:C2",
      "=B1:C2 A2:B2",
      "=A1 + B1",
    ],
    ["4", "5", "6", "", "", "", "", ""],
  ]);
  const snapshot = evaluateSheet(sheet, 21);

  assert.equal(getCellEvaluation(snapshot, 0, 3).display, "5");
  assert.equal(getCellEvaluation(snapshot, 0, 4).display, "#VALUE!");
  assert.equal(getCellEvaluation(snapshot, 0, 5).display, "#NULL!");
  assert.equal(getCellEvaluation(snapshot, 0, 6).display, "5");
  assert.equal(getCellEvaluation(snapshot, 0, 7).display, "3");
  assert.deepEqual(getCellEvaluation(snapshot, 0, 3).dependencies, [cellKey(0, 1), cellKey(0, 2)]);
  assert.deepEqual(getCellEvaluation(snapshot, 0, 5).dependencies, []);
});

test("evaluateSheet supports Excel-style parenthesized reference unions", () => {
  const sheet = createSheet([
    [
      "1",
      "2",
      "3",
      "=SUM((A1:A2,C1:C2))",
      "=SUM((A1:A2,A2:A3))",
      "=(A1,C1)",
      "=INDEX((A1:A2,C1:C2),1)",
      "=SUM((A1:A3,C1:C3) A2:C2)",
    ],
    ["4", "5", "6", "", "", "", "", ""],
    ["7", "8", "9", "", "", "", "", ""],
  ]);
  const snapshot = evaluateSheet(sheet, 22);

  assert.equal(getCellEvaluation(snapshot, 0, 3).display, "14");
  assert.equal(getCellEvaluation(snapshot, 0, 4).display, "16");
  assert.equal(getCellEvaluation(snapshot, 0, 5).display, "#VALUE!");
  assert.equal(getCellEvaluation(snapshot, 0, 6).display, "#VALUE!");
  assert.equal(getCellEvaluation(snapshot, 0, 7).display, "10");
  assert.deepEqual(getCellEvaluation(snapshot, 0, 3).dependencies, [
    cellKey(0, 0),
    cellKey(1, 0),
    cellKey(0, 2),
    cellKey(1, 2),
  ]);
});

test("evaluateWorkbook supports cross-sheet references, ranges, lookups, and cycles", () => {
  const summarySheet = createSheet([
    [
      "=Data!B1+'Data Table'!B1",
      "=SUM('Data Table'!A1:B1)",
      '=XLOOKUP("b",Data!A1:A2,Data!B1:B2)',
      "=Data!C1",
    ],
  ]);
  const dataSheet: WorkbookSheet = {
    ...createSheet([
      ["a", "10", "=Summary!D1"],
      ["b", "20", ""],
    ]),
    id: "sheet-data",
    name: "Data",
  };
  const dataTableSheet: WorkbookSheet = {
    ...createSheet([["1", "2"]]),
    id: "sheet-data-table",
    name: "Data Table",
  };
  const workbook = {
    sheets: [
      {
        ...summarySheet,
        id: "sheet-summary",
        name: "Summary",
      },
      dataSheet,
      dataTableSheet,
    ],
  };
  const snapshots = evaluateWorkbook(workbook, 41);
  const summarySnapshot = snapshots.get("sheet-summary");
  const dataSnapshot = snapshots.get("sheet-data");

  assert.ok(summarySnapshot);
  assert.ok(dataSnapshot);
  assert.equal(getCellEvaluation(summarySnapshot, 0, 0).display, "12");
  assert.equal(getCellEvaluation(summarySnapshot, 0, 1).display, "3");
  assert.equal(getCellEvaluation(summarySnapshot, 0, 2).display, "20");
  assert.equal(getCellEvaluation(summarySnapshot, 0, 3).display, "#CYCLE!");
  assert.equal(getCellEvaluation(dataSnapshot, 0, 2).display, "#CYCLE!");
});

test("evaluateSheet supports core math, logical, and text functions over same-sheet values", () => {
  const sheet = createSheet([
    [
      "1",
      "2",
      "3",
      "text",
      "",
      "=SUM(A1:E1)",
      "=PRODUCT(A1:C1)",
      "=MIN(A1:C1)",
      "=MAX(A1:C1)",
      "=AVERAGE(A1:C1)",
      "=COUNT(A1:E1)",
      "=COUNTA(A1:E1)",
    ],
    [
      "=ABS(-5)",
      "=ROUND(12.345,2)",
      "=INT(3.9)",
      "=MOD(-3,2)",
      "=POWER(2,4)",
      "=SQRT(9)",
      "=TRUE()",
      "=FALSE()",
      "=AND(TRUE,1)",
      "=OR(FALSE,0,2)",
      "=NOT(TRUE)",
      "",
    ],
    [
      '=LEN("abc")',
      '=LEFT("hello",2)',
      '=RIGHT("hello",2)',
      '=MID("hello",2,3)',
      '=TRIM("  a   b  ")',
      '=LOWER("AbC")',
      '=UPPER("AbC")',
      '=CONCAT("a",1,TRUE)',
      '=TEXTJOIN(", ",TRUE,"a","",LOWER("B"))',
      '=VALUE("12.5")',
      "",
      "",
    ],
  ]);
  const snapshot = evaluateSheet(sheet, 23);

  assert.deepEqual(
    [
      getCellEvaluation(snapshot, 0, 5).display,
      getCellEvaluation(snapshot, 0, 6).display,
      getCellEvaluation(snapshot, 0, 7).display,
      getCellEvaluation(snapshot, 0, 8).display,
      getCellEvaluation(snapshot, 0, 9).display,
      getCellEvaluation(snapshot, 0, 10).display,
      getCellEvaluation(snapshot, 0, 11).display,
    ],
    ["6", "6", "1", "3", "2", "3", "4"],
  );
  assert.deepEqual(
    [
      getCellEvaluation(snapshot, 1, 0).display,
      getCellEvaluation(snapshot, 1, 1).display,
      getCellEvaluation(snapshot, 1, 2).display,
      getCellEvaluation(snapshot, 1, 3).display,
      getCellEvaluation(snapshot, 1, 4).display,
      getCellEvaluation(snapshot, 1, 5).display,
      getCellEvaluation(snapshot, 1, 6).display,
      getCellEvaluation(snapshot, 1, 7).display,
      getCellEvaluation(snapshot, 1, 8).display,
      getCellEvaluation(snapshot, 1, 9).display,
      getCellEvaluation(snapshot, 1, 10).display,
    ],
    ["5", "12.35", "3", "1", "16", "3", "TRUE", "FALSE", "TRUE", "TRUE", "FALSE"],
  );
  assert.deepEqual(
    [
      getCellEvaluation(snapshot, 2, 0).display,
      getCellEvaluation(snapshot, 2, 1).display,
      getCellEvaluation(snapshot, 2, 2).display,
      getCellEvaluation(snapshot, 2, 3).display,
      getCellEvaluation(snapshot, 2, 4).display,
      getCellEvaluation(snapshot, 2, 5).display,
      getCellEvaluation(snapshot, 2, 6).display,
      getCellEvaluation(snapshot, 2, 7).display,
      getCellEvaluation(snapshot, 2, 8).display,
      getCellEvaluation(snapshot, 2, 9).display,
    ],
    ["3", "he", "lo", "ell", "a b", "abc", "ABC", "a1TRUE", "a, b", "12.5"],
  );
});

test("evaluateSheet evaluates IF and IFERROR lazily", () => {
  const sheet = createSheet([
    ["4", "0", "  spaced  ", '=IF(B1=0,"zero",A1/B1)', "=IFERROR(A1/B1,99)"],
  ]);
  const snapshot = evaluateSheet(sheet, 29);

  assert.equal(getCellEvaluation(snapshot, 0, 3).display, "zero");
  assert.equal(getCellEvaluation(snapshot, 0, 4).display, "99");
});

test("evaluateSheet supports same-sheet lookup and reference functions", () => {
  const sheet = createSheet([
    [
      "a",
      "10",
      "100",
      '=MATCH("b",A1:A3)',
      '=XLOOKUP("c",A1:A3,B1:B3,"nf")',
      "=INDEX(B1:B3,2)",
      '=CHOOSE(2,"x","y","z")',
    ],
    [
      "b",
      "20",
      "200",
      "=MATCH(15,B1:B3,1)",
      '=XLOOKUP("x",A1:A3,B1:B3,"nf")',
      "=INDEX(A1:C3,2,3)",
      "=ROW(A3)",
    ],
    ["c", "30", "300", "=MATCH(25,B1:B3,-1)", "=ROW()", "=COLUMN()", "=COLUMN(C1)"],
  ]);
  const snapshot = evaluateSheet(sheet, 31);

  assert.deepEqual(
    [
      getCellEvaluation(snapshot, 0, 3).display,
      getCellEvaluation(snapshot, 0, 4).display,
      getCellEvaluation(snapshot, 0, 5).display,
      getCellEvaluation(snapshot, 0, 6).display,
      getCellEvaluation(snapshot, 1, 3).display,
      getCellEvaluation(snapshot, 1, 4).display,
      getCellEvaluation(snapshot, 1, 5).display,
      getCellEvaluation(snapshot, 1, 6).display,
      getCellEvaluation(snapshot, 2, 3).display,
      getCellEvaluation(snapshot, 2, 4).display,
      getCellEvaluation(snapshot, 2, 5).display,
      getCellEvaluation(snapshot, 2, 6).display,
    ],
    ["2", "30", "20", "y", "1", "nf", "200", "3", "3", "3", "6", "3"],
  );
});

test("evaluateSheet supports VLOOKUP with exact default, approximate opt-in, and lookup errors", () => {
  const sheet = createSheet([
    [
      "ID",
      "Name",
      "Score",
      '=VLOOKUP("b",A2:C4,2)',
      '=VLOOKUP("B",A2:C4,3,FALSE)',
      "=VLOOKUP(85,A7:B9,2,TRUE)",
      '=VLOOKUP("x",A2:C4,2)',
      '=VLOOKUP("a",A2:C4,0)',
      '=VLOOKUP("a",A2:C4,4)',
      '=VLOOKUP("missing",A2:C5,2)',
      '=VLOOKUP("ret",A6:C6,2)',
    ],
    ["a", "Alpha", "90"],
    ["b", "Beta", "80"],
    ["c", "Gamma", "70"],
    ["=1/0", "Error row", ""],
    ["ret", "=#N/A", ""],
    ["0", "F", ""],
    ["60", "D", ""],
    ["80", "B", ""],
  ]);
  const snapshot = evaluateSheet(sheet, 37);

  assert.deepEqual(
    [
      getCellEvaluation(snapshot, 0, 3).display,
      getCellEvaluation(snapshot, 0, 4).display,
      getCellEvaluation(snapshot, 0, 5).display,
      getCellEvaluation(snapshot, 0, 6).display,
      getCellEvaluation(snapshot, 0, 7).display,
      getCellEvaluation(snapshot, 0, 8).display,
      getCellEvaluation(snapshot, 0, 9).display,
      getCellEvaluation(snapshot, 0, 10).display,
    ],
    ["Beta", "80", "B", "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#N/A"],
  );
});
