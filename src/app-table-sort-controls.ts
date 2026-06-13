import type {
  WorkbookTableSortDirection,
  WorkbookTableSummary,
  WorkbookTransactionOperation,
} from "./workbook-core";

export interface TableHeaderSortTarget {
  columnIndex: number;
  direction?: WorkbookTableSortDirection;
  rowIndex: number;
  tableId: string;
}

export interface VisibleGridRegion {
  height: number;
  width: number;
  x: number;
  y: number;
}

export function getVisibleTableHeaderSortTargets(
  tables: readonly WorkbookTableSummary[],
  region: VisibleGridRegion | null,
): TableHeaderSortTarget[] {
  if (!region) {
    return [];
  }

  const visibleStartColumn = region.x;
  const visibleEndColumn = region.x + region.width;
  const visibleStartRow = region.y;
  const visibleEndRow = region.y + region.height;
  const targets: TableHeaderSortTarget[] = [];

  for (const table of tables) {
    if (!table.hasHeaderRow) {
      continue;
    }

    const headerRowIndex = table.range.startRow;

    if (headerRowIndex < visibleStartRow || headerRowIndex >= visibleEndRow) {
      continue;
    }

    const startColumn = Math.max(table.range.startColumn, visibleStartColumn);
    const endColumn = Math.min(table.range.startColumn + table.range.columnCount, visibleEndColumn);

    for (let columnIndex = startColumn; columnIndex < endColumn; columnIndex += 1) {
      targets.push({
        columnIndex,
        direction:
          table.sortState?.keys[0]?.columnIndex === columnIndex
            ? table.sortState.keys[0].direction
            : undefined,
        rowIndex: headerRowIndex,
        tableId: table.id,
      });
    }
  }

  return targets;
}

export function createSortTableOperation(
  tableId: string,
  columnIndex: number,
  direction: WorkbookTableSortDirection,
): WorkbookTransactionOperation {
  return {
    keys: [
      {
        columnIndex,
        direction,
      },
    ],
    tableId,
    type: "sortTable",
    valueMode: "display",
  };
}

export function getNextTableSortDirection(
  direction?: WorkbookTableSortDirection,
): WorkbookTableSortDirection {
  return direction === "ascending" ? "descending" : "ascending";
}
