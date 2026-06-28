import type { WorkbookCellStyle, WorkbookTableSummary } from "./workbook-core";

const DEFAULT_HIGHLIGHT_BACKGROUND_COLOR = "#fff1c2";
const DEFAULT_HIGHLIGHT_TEXT_COLOR = "#7c2d12";

export function getTableColumnHighlightStyle(
  tables: readonly WorkbookTableSummary[],
  columnIndex: number,
  rowIndex: number,
  displayValue: string,
): WorkbookCellStyle | undefined {
  const numericValue = parseDisplayNumber(displayValue);

  if (numericValue === null) {
    return undefined;
  }

  for (const table of tables) {
    const headerOffset = table.hasHeaderRow ? 1 : 0;
    const bodyStartRow = table.range.startRow + headerOffset;
    const tableEndRow = table.range.startRow + table.range.rowCount;
    const tableStartColumn = table.range.startColumn;
    const tableEndColumn = table.range.startColumn + table.range.columnCount;

    if (
      rowIndex < bodyStartRow ||
      rowIndex >= tableEndRow ||
      columnIndex < tableStartColumn ||
      columnIndex >= tableEndColumn
    ) {
      continue;
    }

    const rule = table.columnHighlightRules?.find(
      (candidate) => candidate.columnIndex === columnIndex,
    );

    if (!rule || numericValue <= rule.threshold) {
      continue;
    }

    return {
      backgroundColor: rule.backgroundColor ?? DEFAULT_HIGHLIGHT_BACKGROUND_COLOR,
      bold: rule.bold ?? true,
      textColor: rule.textColor ?? DEFAULT_HIGHLIGHT_TEXT_COLOR,
    };
  }

  return undefined;
}

function parseDisplayNumber(value: string): number | null {
  const normalizedValue = value.trim().replace(/,/gu, "");

  if (normalizedValue.length === 0) {
    return null;
  }

  const numericText = normalizedValue.endsWith("%")
    ? normalizedValue.slice(0, -1)
    : normalizedValue;

  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/u.test(numericText)) {
    return null;
  }

  const numericValue = Number.parseFloat(numericText);

  return Number.isFinite(numericValue) ? numericValue : null;
}
