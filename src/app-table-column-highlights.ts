import type {
  WorkbookCellStyle,
  WorkbookTableColumnHighlightRule,
  WorkbookTableSummary,
} from "./workbook-core";

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

    let highlightStyle: WorkbookCellStyle | undefined;

    for (const rule of table.columnHighlightRules ?? []) {
      if (
        rule.columnIndex !== columnIndex ||
        !tableColumnHighlightRuleMatches(rule, numericValue)
      ) {
        continue;
      }

      highlightStyle = mergeTableColumnHighlightStyle(highlightStyle, rule);
    }

    if (highlightStyle) {
      return highlightStyle;
    }
  }

  return undefined;
}

function tableColumnHighlightRuleMatches(
  rule: WorkbookTableColumnHighlightRule,
  numericValue: number,
): boolean {
  switch (rule.thresholdType ?? "higher") {
    case "equal":
      return numericValue === rule.threshold;
    case "higher":
      return numericValue > rule.threshold;
    case "higherOrEqual":
      return numericValue >= rule.threshold;
    case "lower":
      return numericValue < rule.threshold;
    case "lowerOrEqual":
      return numericValue <= rule.threshold;
  }
}

function mergeTableColumnHighlightStyle(
  current: WorkbookCellStyle | undefined,
  rule: WorkbookTableColumnHighlightRule,
): WorkbookCellStyle {
  if (!current) {
    return {
      backgroundColor: rule.backgroundColor ?? DEFAULT_HIGHLIGHT_BACKGROUND_COLOR,
      bold: rule.bold ?? true,
      textColor: rule.textColor ?? DEFAULT_HIGHLIGHT_TEXT_COLOR,
    };
  }

  return {
    ...current,
    ...(rule.backgroundColor ? { backgroundColor: rule.backgroundColor } : {}),
    ...(rule.bold !== undefined ? { bold: rule.bold } : {}),
    ...(rule.textColor ? { textColor: rule.textColor } : {}),
  };
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
