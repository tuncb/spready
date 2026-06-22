import type { SheetDisplayRangeResult, UsedRangeResult, WorkbookSummary } from "./workbook-core";

export interface WorkbookConsoleOutputSheet {
  displayRange?: SheetDisplayRangeResult;
  sheetId: string;
  sheetName: string;
  usedRange: UsedRangeResult;
}

export function formatWorkbookConsoleOutput(
  summary: WorkbookSummary,
  sheets: WorkbookConsoleOutputSheet[],
): string {
  const lines = [
    `Workbook: ${summary.documentFilePath ?? "(unsaved workbook)"}`,
    `Sheets: ${summary.sheets.length}`,
    "",
  ];

  for (const sheet of sheets) {
    lines.push(`Sheet: ${sheet.sheetName} (${formatUsedRangeLabel(sheet.usedRange)})`);

    if (
      !sheet.displayRange ||
      sheet.displayRange.rowCount === 0 ||
      sheet.displayRange.columnCount === 0
    ) {
      lines.push("  (empty)");
      lines.push("");
      continue;
    }

    for (const line of formatGrid(
      sheet.displayRange.values,
      sheet.displayRange.startRow,
      sheet.displayRange.startColumn,
    )) {
      lines.push(`  ${line}`);
    }

    lines.push("");
  }

  return `${lines.join("\n").trimEnd()}\n`;
}

function formatUsedRangeLabel(usedRange: UsedRangeResult) {
  if (usedRange.rowCount === 0 || usedRange.columnCount === 0) {
    return "empty";
  }

  const firstCell = formatCellAddress(usedRange.startRow, usedRange.startColumn);
  const lastCell = formatCellAddress(
    usedRange.startRow + usedRange.rowCount - 1,
    usedRange.startColumn + usedRange.columnCount - 1,
  );

  return firstCell === lastCell ? firstCell : `${firstCell}:${lastCell}`;
}

function formatGrid(values: string[][], startRow: number, startColumn: number) {
  const rowCount = values.length;
  const columnCount = Math.max(0, ...values.map((row) => row.length));
  const rowLabels = Array.from({ length: rowCount }, (_value, rowOffset) =>
    String(startRow + rowOffset + 1),
  );
  const columnLabels = Array.from({ length: columnCount }, (_value, columnOffset) =>
    formatColumnLabel(startColumn + columnOffset),
  );
  const columnWidths = columnLabels.map((label, columnOffset) =>
    Math.max(label.length, ...values.map((row) => formatCellValue(row[columnOffset] ?? "").length)),
  );
  const rowLabelWidth = Math.max(1, ...rowLabels.map((label) => label.length));
  const lines = [
    [
      "".padStart(rowLabelWidth),
      ...columnLabels.map((label, index) => label.padEnd(columnWidths[index])),
    ]
      .join("  ")
      .trimEnd(),
  ];

  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    lines.push(
      [
        rowLabels[rowOffset].padStart(rowLabelWidth),
        ...Array.from({ length: columnCount }, (_value, columnOffset) =>
          formatCellValue(values[rowOffset]?.[columnOffset] ?? "").padEnd(
            columnWidths[columnOffset],
          ),
        ),
      ]
        .join("  ")
        .trimEnd(),
    );
  }

  return lines;
}

function formatCellValue(value: string) {
  return value.replaceAll("\r", "\\r").replaceAll("\n", "\\n").replaceAll("\t", "\\t");
}

function formatCellAddress(rowIndex: number, columnIndex: number) {
  return `${formatColumnLabel(columnIndex)}${rowIndex + 1}`;
}

function formatColumnLabel(columnIndex: number) {
  let value = columnIndex + 1;
  let label = "";

  while (value > 0) {
    value -= 1;
    label = `${String.fromCharCode(65 + (value % 26))}${label}`;
    value = Math.floor(value / 26);
  }

  return label;
}
