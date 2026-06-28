import type { Item } from "@glideapps/glide-data-grid";

import { getColumnTitle, type SheetRangeRequest, type WorkbookTableSummary } from "./workbook-core";

export interface TableRowHintTarget {
  activeColumnIndex: number;
  rowIndex: number;
  table: WorkbookTableSummary;
  tableRowNumber: number;
}

export interface TableRowHintRangeRequests {
  header: SheetRangeRequest;
  row: SheetRangeRequest;
}

export interface TableRowHintItem {
  columnIndex: number;
  isActive: boolean;
  label: string;
  value: string;
}

export function getTableRowHintTarget(
  tables: readonly WorkbookTableSummary[],
  cell: Item | null,
): TableRowHintTarget | null {
  if (!cell) {
    return null;
  }

  const [columnIndex, rowIndex] = cell;

  for (const table of tables) {
    if (!table.hasHeaderRow) {
      continue;
    }

    const headerRowIndex = table.range.startRow;
    const tableEndRow = table.range.startRow + table.range.rowCount;
    const tableStartColumn = table.range.startColumn;
    const tableEndColumn = table.range.startColumn + table.range.columnCount;

    if (
      rowIndex <= headerRowIndex ||
      rowIndex >= tableEndRow ||
      columnIndex < tableStartColumn ||
      columnIndex >= tableEndColumn
    ) {
      continue;
    }

    return {
      activeColumnIndex: columnIndex,
      rowIndex,
      table,
      tableRowNumber: rowIndex - headerRowIndex,
    };
  }

  return null;
}

export function buildTableRowHintRangeRequests(
  target: TableRowHintTarget,
): TableRowHintRangeRequests {
  const baseRange = {
    columnCount: target.table.range.columnCount,
    sheetId: target.table.sheetId,
    startColumn: target.table.range.startColumn,
  };

  return {
    header: {
      ...baseRange,
      rowCount: 1,
      startRow: target.table.range.startRow,
    },
    row: {
      ...baseRange,
      rowCount: 1,
      startRow: target.rowIndex,
    },
  };
}

export function buildTableRowHintItems(
  target: TableRowHintTarget,
  headerValues: readonly string[],
  rowValues: readonly string[],
): TableRowHintItem[] {
  return Array.from({ length: target.table.range.columnCount }, (_, columnOffset) => {
    const columnIndex = target.table.range.startColumn + columnOffset;
    const headerValue = headerValues[columnOffset]?.trim();

    return {
      columnIndex,
      isActive: columnIndex === target.activeColumnIndex,
      label: headerValue || getColumnTitle(columnIndex),
      value: rowValues[columnOffset] ?? "",
    };
  });
}

export function getTableRowHintTargetKey(
  target: TableRowHintTarget,
  workbookVersion: number,
): string {
  return `${target.table.sheetId}:${target.table.id}:${target.rowIndex}:${target.activeColumnIndex}:${workbookVersion}`;
}

export function getTableRowHintTableKey(target: TableRowHintTarget): string {
  return `${target.table.sheetId}:${target.table.id}`;
}
