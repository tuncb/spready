import * as z from "zod/v4";

import {
  assertWorkbookSheetNamesAreUnique,
  cloneWorkbookCellStyle,
  createSheet,
  MAX_COLUMN_WIDTH,
  MIN_CHART_LAYOUT_HEIGHT,
  MIN_CHART_LAYOUT_WIDTH,
  MIN_COLUMN_WIDTH,
  syncSheetIdSequence,
  type WorkbookCellStyle,
  type WorkbookChart,
  type WorkbookSheet,
  type WorkbookState,
  type WorkbookTable,
} from "./workbook-core";

export const WORKBOOK_DOCUMENT_EXTENSION = ".spready";
export const WORKBOOK_DOCUMENT_FORMAT = "spready-workbook";
export const WORKBOOK_DOCUMENT_VERSION = 6;

export interface WorkbookDocumentCell {
  column: number;
  row: number;
  value: string;
}

export interface WorkbookDocumentCellStyle {
  column: number;
  row: number;
  style: WorkbookCellStyle;
}

export interface WorkbookDocumentSheet {
  cells: WorkbookDocumentCell[];
  columnCount: number;
  columnWidths?: Record<string, number>;
  id: string;
  metadata?: {
    sourceFilePath?: string;
  };
  name: string;
  rowCount: number;
  styles: WorkbookDocumentCellStyle[];
}

export type WorkbookDocumentChart = WorkbookChart;
export type WorkbookDocumentTable = WorkbookTable;

export interface WorkbookDocument {
  format: typeof WORKBOOK_DOCUMENT_FORMAT;
  formatVersion: 4 | 5 | typeof WORKBOOK_DOCUMENT_VERSION;
  workbook: {
    activeSheetId: string;
    charts: WorkbookDocumentChart[];
    nextChartNumber: number;
    nextSheetNumber: number;
    nextTableNumber?: number;
    sheets: WorkbookDocumentSheet[];
    tables?: WorkbookDocumentTable[];
  };
}

const workbookDocumentCellSchema = z.object({
  column: z.int().min(0),
  row: z.int().min(0),
  value: z.string(),
});

const workbookDocumentCellStyleValueSchema = z
  .object({
    backgroundColor: z.string().min(1).optional(),
    bold: z.boolean().optional(),
    fontFamily: z.string().min(1).optional(),
    fontSize: z.number().min(6).optional(),
    horizontalAlign: z.enum(["center", "left", "right"]).optional(),
    italic: z.boolean().optional(),
    numberFormat: z
      .discriminatedUnion("type", [
        z.object({
          type: z.literal("general"),
        }),
        z
          .object({
            decimalPlaces: z.int().min(0).max(20).optional(),
            significantDigits: z.int().min(1).max(21).optional(),
            type: z.literal("number"),
            useGrouping: z.boolean().optional(),
          })
          .refine(
            (format) =>
              !(format.decimalPlaces !== undefined && format.significantDigits !== undefined),
            {
              message: "Number format cannot set both decimalPlaces and significantDigits.",
            },
          ),
        z
          .object({
            decimalPlaces: z.int().min(0).max(20).optional(),
            significantDigits: z.int().min(1).max(21).optional(),
            type: z.literal("percent"),
            useGrouping: z.boolean().optional(),
          })
          .refine(
            (format) =>
              !(format.decimalPlaces !== undefined && format.significantDigits !== undefined),
            {
              message: "Number format cannot set both decimalPlaces and significantDigits.",
            },
          ),
        z.object({
          significantDigits: z.int().min(1).max(21).optional(),
          type: z.literal("scientific"),
        }),
      ])
      .optional(),
    textColor: z.string().min(1).optional(),
    wrapText: z.boolean().optional(),
  })
  .refine((style) => Object.keys(style).length > 0, {
    message: "Cell style must contain at least one style property.",
  });

const workbookDocumentCellStyleSchema = z.object({
  column: z.int().min(0),
  row: z.int().min(0),
  style: workbookDocumentCellStyleValueSchema,
});

const workbookDocumentSheetSchema = z.object({
  cells: z.array(workbookDocumentCellSchema),
  columnCount: z.int().min(1),
  columnWidths: z
    .record(z.string(), z.number().min(MIN_COLUMN_WIDTH).max(MAX_COLUMN_WIDTH))
    .optional(),
  id: z.string().min(1),
  metadata: z
    .object({
      sourceFilePath: z.string().min(1).optional(),
    })
    .optional(),
  name: z.string().min(1),
  rowCount: z.int().min(1),
  styles: z.array(workbookDocumentCellStyleSchema),
});

const workbookDocumentChartRangeSchema = z.object({
  columnCount: z.int().min(0),
  rowCount: z.int().min(0),
  sheetId: z.string().min(1),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
});

const workbookDocumentChartSourceSchema = z.object({
  range: workbookDocumentChartRangeSchema,
  seriesLayoutBy: z.enum(["column", "row"]),
  sourceHeader: z.boolean(),
});

const workbookDocumentChartLayoutSchema = z.object({
  height: z.number().min(MIN_CHART_LAYOUT_HEIGHT),
  offsetX: z.number().min(0),
  offsetY: z.number().min(0),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
  width: z.number().min(MIN_CHART_LAYOUT_WIDTH),
  zIndex: z.int().min(0),
});

const workbookDocumentCartesianChartSpecSchema = z.object({
  categoryDimension: z.int().min(0),
  chartType: z.enum(["bar", "line", "area", "scatter"]),
  family: z.literal("cartesian"),
  smooth: z.boolean().optional(),
  source: workbookDocumentChartSourceSchema,
  stacked: z.boolean().optional(),
  valueDimensions: z.array(z.int().min(0)),
});

const workbookDocumentPieChartSpecSchema = z.object({
  chartType: z.literal("pie"),
  family: z.literal("pie"),
  nameDimension: z.int().min(0),
  source: workbookDocumentChartSourceSchema,
  valueDimension: z.int().min(0),
});

const workbookDocumentChartSchema = z.object({
  id: z.string().min(1),
  layout: workbookDocumentChartLayoutSchema,
  name: z.string().min(1),
  sheetId: z.string().min(1),
  spec: z.discriminatedUnion("family", [
    workbookDocumentCartesianChartSpecSchema,
    workbookDocumentPieChartSpecSchema,
  ]),
});

const workbookDocumentTableRangeSchema = z.object({
  columnCount: z.int().min(1),
  rowCount: z.int().min(1),
  sheetId: z.string().min(1),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
});

const workbookDocumentTableSortStateSchema = z.object({
  keys: z
    .array(
      z.object({
        columnIndex: z.int().min(0),
        direction: z.enum(["ascending", "descending"]),
      }),
    )
    .min(1),
  valueMode: z.enum(["display", "raw"]),
});

const workbookDocumentTableColumnHighlightRuleSchema = z.object({
  backgroundColor: z.string().min(1).optional(),
  bold: z.boolean().optional(),
  columnIndex: z.int().min(0),
  textColor: z.string().min(1).optional(),
  threshold: z.number(),
});

const workbookDocumentTableSchema = z.object({
  columnHighlightRules: z.array(workbookDocumentTableColumnHighlightRuleSchema).optional(),
  hasHeaderRow: z.boolean(),
  id: z.string().min(1),
  name: z.string().min(1),
  range: workbookDocumentTableRangeSchema,
  sortState: workbookDocumentTableSortStateSchema.optional(),
});

const workbookDocumentSchema = z.object({
  format: z.literal(WORKBOOK_DOCUMENT_FORMAT),
  formatVersion: z.union([z.literal(4), z.literal(5), z.literal(WORKBOOK_DOCUMENT_VERSION)]),
  workbook: z.object({
    activeSheetId: z.string().min(1),
    charts: z.array(workbookDocumentChartSchema),
    nextChartNumber: z.int().min(1),
    nextSheetNumber: z.int().min(1),
    nextTableNumber: z.int().min(1).optional(),
    sheets: z.array(workbookDocumentSheetSchema).min(1),
    tables: z.array(workbookDocumentTableSchema).optional(),
  }),
});

export function createWorkbookDocument(workbook: WorkbookState): WorkbookDocument {
  return {
    format: WORKBOOK_DOCUMENT_FORMAT,
    formatVersion: WORKBOOK_DOCUMENT_VERSION,
    workbook: {
      activeSheetId: workbook.activeSheetId,
      charts: workbook.charts.map((chart) => ({ ...chart })),
      nextChartNumber: workbook.nextChartNumber,
      nextSheetNumber: workbook.nextSheetNumber,
      nextTableNumber: workbook.nextTableNumber,
      sheets: workbook.sheets.map((sheet) => createWorkbookDocumentSheet(sheet)),
      tables: workbook.tables.map((table) => restoreWorkbookTable(table)),
    },
  };
}

export function serializeWorkbookDocument(workbook: WorkbookState): string {
  return JSON.stringify(createWorkbookDocument(workbook), null, 2);
}

export function parseWorkbookDocument(content: string): WorkbookState {
  let parsedJson: unknown;

  try {
    parsedJson = JSON.parse(content) as unknown;
  } catch {
    throw new Error("Workbook file must contain valid JSON.");
  }

  const document = parseWorkbookDocumentJson(parsedJson);
  const chartIds = new Set<string>();
  const sheetIds = new Set<string>();
  const tableIds = new Set<string>();
  const tableNames = new Map<string, string>();

  assertWorkbookSheetNamesAreUnique({ sheets: document.workbook.sheets });

  for (const sheet of document.workbook.sheets) {
    if (sheetIds.has(sheet.id)) {
      throw new Error(`Workbook file contains a duplicate sheet id "${sheet.id}".`);
    }

    sheetIds.add(sheet.id);
  }

  for (const chart of document.workbook.charts) {
    if (chartIds.has(chart.id)) {
      throw new Error(`Workbook file contains a duplicate chart id "${chart.id}".`);
    }

    chartIds.add(chart.id);
  }

  for (const table of document.workbook.tables ?? []) {
    if (tableIds.has(table.id)) {
      throw new Error(`Workbook file contains a duplicate table id "${table.id}".`);
    }

    tableIds.add(table.id);

    const tableNameKey = table.name.trim().toLowerCase();
    const existingName = tableNames.get(tableNameKey);

    if (existingName) {
      throw new Error(
        `Workbook file contains duplicate table names "${existingName}" and "${table.name}".`,
      );
    }

    tableNames.set(tableNameKey, table.name);

    const sheet = document.workbook.sheets.find((entry) => entry.id === table.range.sheetId);

    if (!sheet) {
      throw new Error(`Workbook file table "${table.id}" references missing sheet.`);
    }

    if (
      table.range.startRow + table.range.rowCount > sheet.rowCount ||
      table.range.startColumn + table.range.columnCount > sheet.columnCount
    ) {
      throw new Error(`Workbook file table "${table.id}" is outside sheet bounds.`);
    }

    for (const key of table.sortState?.keys ?? []) {
      if (
        key.columnIndex < table.range.startColumn ||
        key.columnIndex >= table.range.startColumn + table.range.columnCount
      ) {
        throw new Error(`Workbook file table "${table.id}" has an out-of-bounds sort key.`);
      }
    }

    const highlightColumns = new Set<number>();

    for (const rule of table.columnHighlightRules ?? []) {
      if (!Number.isFinite(rule.threshold)) {
        throw new Error(`Workbook file table "${table.id}" has a non-finite highlight threshold.`);
      }

      if (
        rule.columnIndex < table.range.startColumn ||
        rule.columnIndex >= table.range.startColumn + table.range.columnCount
      ) {
        throw new Error(`Workbook file table "${table.id}" has an out-of-bounds highlight column.`);
      }

      if (highlightColumns.has(rule.columnIndex)) {
        throw new Error(
          `Workbook file table "${table.id}" has duplicate highlight rules for column ${rule.columnIndex}.`,
        );
      }

      highlightColumns.add(rule.columnIndex);
    }
  }

  for (let leftIndex = 0; leftIndex < (document.workbook.tables ?? []).length; leftIndex += 1) {
    for (
      let rightIndex = leftIndex + 1;
      rightIndex < (document.workbook.tables ?? []).length;
      rightIndex += 1
    ) {
      const left = document.workbook.tables?.[leftIndex];
      const right = document.workbook.tables?.[rightIndex];

      if (left && right && tableRangesOverlap(left.range, right.range)) {
        throw new Error(`Workbook file tables "${left.id}" and "${right.id}" overlap.`);
      }
    }
  }

  if (!sheetIds.has(document.workbook.activeSheetId)) {
    throw new Error(
      `Workbook file references missing active sheet "${document.workbook.activeSheetId}".`,
    );
  }

  const workbook: WorkbookState = {
    activeSheetId: document.workbook.activeSheetId,
    charts: document.workbook.charts.map((chart) => restoreWorkbookChart(chart)),
    hasUnsavedChanges: false,
    nextChartNumber: document.workbook.nextChartNumber,
    nextSheetNumber: document.workbook.nextSheetNumber,
    nextTableNumber: document.workbook.nextTableNumber ?? 1,
    sheets: document.workbook.sheets.map((sheet) => restoreWorkbookSheet(sheet)),
    tables: (document.workbook.tables ?? []).map((table) => restoreWorkbookTable(table)),
    version: 0,
  };

  syncSheetIdSequence(workbook);
  workbook.nextTableNumber = Math.max(
    workbook.nextTableNumber,
    getNextTableNumber(workbook.tables),
  );

  return workbook;
}

function createWorkbookDocumentSheet(sheet: WorkbookSheet): WorkbookDocumentSheet {
  const cells: WorkbookDocumentCell[] = [];
  const styles: WorkbookDocumentCellStyle[] = [];

  for (let rowIndex = 0; rowIndex < sheet.cells.length; rowIndex += 1) {
    const row = sheet.cells[rowIndex];

    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      const value = row[columnIndex] ?? "";

      if (value === "") {
        continue;
      }

      cells.push({
        column: columnIndex,
        row: rowIndex,
        value,
      });
    }
  }

  for (const [key, style] of Object.entries(sheet.cellStyles)) {
    const [rowText, columnText] = key.split(":");
    const row = Number.parseInt(rowText, 10);
    const column = Number.parseInt(columnText, 10);

    if (
      !Number.isInteger(row) ||
      !Number.isInteger(column) ||
      row < 0 ||
      column < 0 ||
      row >= sheet.cells.length ||
      column >= (sheet.cells[0]?.length ?? 0)
    ) {
      continue;
    }

    const normalizedStyle = cloneWorkbookCellStyle(style);

    if (Object.keys(normalizedStyle).length === 0) {
      continue;
    }

    styles.push({
      column,
      row,
      style: normalizedStyle,
    });
  }

  return {
    cells,
    columnCount: Math.max(1, sheet.cells[0]?.length ?? 0),
    ...(Object.keys(sheet.columnWidths).length > 0
      ? {
          columnWidths: {
            ...sheet.columnWidths,
          },
        }
      : {}),
    id: sheet.id,
    ...(sheet.sourceFilePath
      ? {
          metadata: {
            sourceFilePath: sheet.sourceFilePath,
          },
        }
      : {}),
    name: sheet.name,
    rowCount: Math.max(1, sheet.cells.length),
    styles,
  };
}

function formatValidationError(error: z.ZodError): string {
  const firstIssue = error.issues[0];

  if (!firstIssue) {
    return "Workbook file is invalid.";
  }

  const path = firstIssue.path.length > 0 ? firstIssue.path.join(".") : "root";

  return `Workbook file is invalid at "${path}": ${firstIssue.message}`;
}

function parseWorkbookDocumentJson(parsedJson: unknown): WorkbookDocument {
  try {
    return workbookDocumentSchema.parse(parsedJson);
  } catch (error) {
    if (error instanceof z.ZodError) {
      throw new Error(formatValidationError(error));
    }

    throw error;
  }
}

function restoreWorkbookSheet(sheet: WorkbookDocumentSheet): WorkbookSheet {
  const cells = createSheet(sheet.rowCount, sheet.columnCount);
  const cellStyles: Record<string, WorkbookCellStyle> = {};
  const columnWidths: Record<string, number> = {};
  const occupiedCellKeys = new Set<string>();
  const styledCellKeys = new Set<string>();

  for (const cell of sheet.cells) {
    if (cell.row >= sheet.rowCount || cell.column >= sheet.columnCount) {
      throw new Error(
        `Workbook file contains out-of-bounds cell ${cell.row}:${cell.column} in sheet "${sheet.id}".`,
      );
    }

    const cellKey = `${cell.row}:${cell.column}`;

    if (occupiedCellKeys.has(cellKey)) {
      throw new Error(
        `Workbook file contains a duplicate cell entry for ${cell.row}:${cell.column} in sheet "${sheet.id}".`,
      );
    }

    occupiedCellKeys.add(cellKey);

    if (cell.value !== "") {
      cells[cell.row][cell.column] = cell.value;
    }
  }

  for (const cellStyle of sheet.styles) {
    if (cellStyle.row >= sheet.rowCount || cellStyle.column >= sheet.columnCount) {
      throw new Error(
        `Workbook file contains out-of-bounds style ${cellStyle.row}:${cellStyle.column} in sheet "${sheet.id}".`,
      );
    }

    const cellKey = `${cellStyle.row}:${cellStyle.column}`;

    if (styledCellKeys.has(cellKey)) {
      throw new Error(
        `Workbook file contains a duplicate style entry for ${cellStyle.row}:${cellStyle.column} in sheet "${sheet.id}".`,
      );
    }

    styledCellKeys.add(cellKey);
    cellStyles[cellKey] = cloneWorkbookCellStyle(cellStyle.style);
  }

  for (const [key, width] of Object.entries(sheet.columnWidths ?? {})) {
    const column = Number(key);

    if (!Number.isInteger(column) || column < 0 || column >= sheet.columnCount) {
      throw new Error(
        `Workbook file contains out-of-bounds column width ${key} in sheet "${sheet.id}".`,
      );
    }

    columnWidths[key] = Math.round(width);
  }

  return {
    cells,
    cellStyles,
    columnWidths,
    id: sheet.id,
    name: sheet.name,
    sourceFilePath: sheet.metadata?.sourceFilePath,
  };
}

function restoreWorkbookChart(chart: WorkbookDocumentChart): WorkbookChart {
  return {
    id: chart.id,
    layout: {
      ...chart.layout,
    },
    name: chart.name,
    sheetId: chart.sheetId,
    spec: {
      ...chart.spec,
      source: {
        ...chart.spec.source,
        range: {
          ...chart.spec.source.range,
        },
      },
    },
  };
}

function restoreWorkbookTable(table: WorkbookDocumentTable): WorkbookTable {
  return {
    ...(table.columnHighlightRules
      ? { columnHighlightRules: table.columnHighlightRules.map((rule) => ({ ...rule })) }
      : {}),
    hasHeaderRow: table.hasHeaderRow,
    id: table.id,
    name: table.name,
    range: {
      ...table.range,
    },
    ...(table.sortState
      ? {
          sortState: {
            keys: table.sortState.keys.map((key) => ({ ...key })),
            valueMode: table.sortState.valueMode,
          },
        }
      : {}),
  };
}

function tableRangesOverlap(
  left: WorkbookDocumentTable["range"],
  right: WorkbookDocumentTable["range"],
): boolean {
  if (left.sheetId !== right.sheetId) {
    return false;
  }

  return (
    left.startRow < right.startRow + right.rowCount &&
    right.startRow < left.startRow + left.rowCount &&
    left.startColumn < right.startColumn + right.columnCount &&
    right.startColumn < left.startColumn + left.columnCount
  );
}

function getNextTableNumber(tables: WorkbookTable[]): number {
  return (
    Math.max(
      0,
      ...tables.map((table) => {
        const match = /^table-(\d+)$/u.exec(table.id);
        return match ? Number.parseInt(match[1], 10) : 0;
      }),
    ) + 1
  );
}
