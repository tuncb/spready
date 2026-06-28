import type { SheetRangeRequest, WorkbookTableSummary } from "./workbook-core";

export interface VisibleGridRegion {
  height: number;
  width: number;
  x: number;
  y: number;
}

export function getStickyTableHeader(
  tables: readonly WorkbookTableSummary[],
  region: VisibleGridRegion | null,
): WorkbookTableSummary | null {
  if (!region) {
    return null;
  }

  const visibleStartColumn = region.x;
  const visibleEndColumn = region.x + region.width;

  for (const table of tables) {
    if (!table.hasHeaderRow) {
      continue;
    }

    const tableStartRow = table.range.startRow;
    const tableBodyStartRow = tableStartRow + 1;
    const tableEndRow = table.range.startRow + table.range.rowCount;
    const tableStartColumn = table.range.startColumn;
    const tableEndColumn = table.range.startColumn + table.range.columnCount;

    if (region.y < tableBodyStartRow || region.y >= tableEndRow) {
      continue;
    }

    if (visibleEndColumn <= tableStartColumn || visibleStartColumn >= tableEndColumn) {
      continue;
    }

    return table;
  }

  return null;
}

export function buildStickyTableHeaderRangeRequest(
  table: WorkbookTableSummary,
  region: VisibleGridRegion | null,
): SheetRangeRequest | null {
  if (!region) {
    return null;
  }

  const startColumn = Math.max(table.range.startColumn, region.x);
  const endColumn = Math.min(
    table.range.startColumn + table.range.columnCount,
    region.x + region.width,
  );

  if (startColumn >= endColumn) {
    return null;
  }

  return {
    columnCount: endColumn - startColumn,
    rowCount: 1,
    sheetId: table.sheetId,
    startColumn,
    startRow: table.range.startRow,
  };
}

export function getTableHeaderCacheKey(
  sheetId: string,
  tableId: string,
  workbookVersion: number,
): string {
  return `${sheetId}:${tableId}:${workbookVersion}`;
}
