import * as z from "zod/v4";

import {
  createSheet,
  syncSheetIdSequence,
  type WorkbookSheet,
  type WorkbookState,
} from "./workbook-core";

export const WORKBOOK_DOCUMENT_EXTENSION = ".spready";
export const WORKBOOK_DOCUMENT_FORMAT = "spready-workbook";
export const WORKBOOK_DOCUMENT_VERSION = 1;

export interface WorkbookDocumentCell {
  column: number;
  row: number;
  value: string;
}

export interface WorkbookDocumentSheet {
  cells: WorkbookDocumentCell[];
  columnCount: number;
  id: string;
  metadata?: {
    sourceFilePath?: string;
  };
  name: string;
  rowCount: number;
}

export interface WorkbookDocument {
  format: typeof WORKBOOK_DOCUMENT_FORMAT;
  formatVersion: typeof WORKBOOK_DOCUMENT_VERSION;
  workbook: {
    activeSheetId: string;
    nextSheetNumber: number;
    sheets: WorkbookDocumentSheet[];
  };
}

const workbookDocumentCellSchema = z.object({
  column: z.int().min(0),
  row: z.int().min(0),
  value: z.string(),
});

const workbookDocumentSheetSchema = z.object({
  cells: z.array(workbookDocumentCellSchema),
  columnCount: z.int().min(1),
  id: z.string().min(1),
  metadata: z
    .object({
      sourceFilePath: z.string().min(1).optional(),
    })
    .optional(),
  name: z.string().min(1),
  rowCount: z.int().min(1),
});

const workbookDocumentSchema = z.object({
  format: z.literal(WORKBOOK_DOCUMENT_FORMAT),
  formatVersion: z.literal(WORKBOOK_DOCUMENT_VERSION),
  workbook: z.object({
    activeSheetId: z.string().min(1),
    nextSheetNumber: z.int().min(1),
    sheets: z.array(workbookDocumentSheetSchema).min(1),
  }),
});

export function createWorkbookDocument(
  workbook: WorkbookState,
): WorkbookDocument {
  return {
    format: WORKBOOK_DOCUMENT_FORMAT,
    formatVersion: WORKBOOK_DOCUMENT_VERSION,
    workbook: {
      activeSheetId: workbook.activeSheetId,
      nextSheetNumber: workbook.nextSheetNumber,
      sheets: workbook.sheets.map((sheet) => createWorkbookDocumentSheet(sheet)),
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
  const sheetIds = new Set<string>();

  for (const sheet of document.workbook.sheets) {
    if (sheetIds.has(sheet.id)) {
      throw new Error(`Workbook file contains a duplicate sheet id "${sheet.id}".`);
    }

    sheetIds.add(sheet.id);
  }

  if (!sheetIds.has(document.workbook.activeSheetId)) {
    throw new Error(
      `Workbook file references missing active sheet "${document.workbook.activeSheetId}".`,
    );
  }

  const workbook: WorkbookState = {
    activeSheetId: document.workbook.activeSheetId,
    nextSheetNumber: document.workbook.nextSheetNumber,
    sheets: document.workbook.sheets.map((sheet) =>
      restoreWorkbookSheet(sheet),
    ),
    version: 0,
  };

  syncSheetIdSequence(workbook);

  return workbook;
}

function createWorkbookDocumentSheet(sheet: WorkbookSheet): WorkbookDocumentSheet {
  const cells: WorkbookDocumentCell[] = [];

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

  return {
    cells,
    columnCount: Math.max(1, sheet.cells[0]?.length ?? 0),
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
  const occupiedCellKeys = new Set<string>();

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

  return {
    cells,
    id: sheet.id,
    name: sheet.name,
    sourceFilePath: sheet.metadata?.sourceFilePath,
  };
}
