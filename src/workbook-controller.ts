import { promises as fs } from "node:fs";
import path from "node:path";
import { EventEmitter } from "node:events";

import {
  applyWorkbookTransaction,
  createWorkbookState,
  getSheetColumnCount,
  type CsvFileOperationResult,
  type ExportCsvFileRequest,
  getWorkbookSheet,
  getSheetCsv,
  getSheetRange,
  type CellDataRequest,
  type CellDataResult,
  getSheetUsedRange,
  getWorkbookSummary,
  getSheetRowCount,
  type ApplyTransactionRequest,
  type ApplyTransactionResult,
  type ImportCsvFileRequest,
  type OpenWorkbookFileRequest,
  type SaveWorkbookFileRequest,
  type SheetRangeRequest,
  type SheetDisplayRangeResult,
  type SheetRangeResult,
  type UsedRangeResult,
  type WorkbookFileOperationResult,
  type WorkbookState,
  type WorkbookSummary,
} from "./workbook-core";
import {
  evaluateSheet,
  getCellEvaluation,
  type SheetEvaluationSnapshot,
} from "./formula-engine";
import {
  parseWorkbookDocument,
  serializeWorkbookDocument,
  WORKBOOK_DOCUMENT_EXTENSION,
} from "./workbook-document";

export class WorkbookController extends EventEmitter {
  #state: WorkbookState = createWorkbookState();
  #sheetEvaluationSnapshots = new Map<string, SheetEvaluationSnapshot>();

  getSummary(): WorkbookSummary {
    return getWorkbookSummary(this.#state);
  }

  getSheetCsv(sheetId?: string): string {
    return getSheetCsv(this.#state, sheetId);
  }

  getSheetRange(request: SheetRangeRequest): SheetRangeResult {
    return getSheetRange(this.#state, request);
  }

  getSheetDisplayRange(request: SheetRangeRequest): SheetDisplayRangeResult {
    const rawRange = getSheetRange(this.#state, request);
    const snapshot = this.#getEvaluationSnapshot(rawRange.sheetId);

    return {
      ...rawRange,
      values: Array.from({ length: rawRange.rowCount }, (_, rowOffset) =>
        Array.from({ length: rawRange.columnCount }, (_, columnOffset) => {
          return getCellEvaluation(
            snapshot,
            rawRange.startRow + rowOffset,
            rawRange.startColumn + columnOffset,
          ).display;
        }),
      ),
    };
  }

  getCellData(request: CellDataRequest): CellDataResult {
    const sheet = getWorkbookSheet(this.#state, request.sheetId);

    assertCellIndex(request.rowIndex, getSheetRowCount(sheet), "Row");
    assertCellIndex(request.columnIndex, getSheetColumnCount(sheet), "Column");

    const evaluation = getCellEvaluation(
      this.#getEvaluationSnapshot(sheet.id),
      request.rowIndex,
      request.columnIndex,
    );

    return {
      columnIndex: request.columnIndex,
      display: evaluation.display,
      input: evaluation.input,
      isFormula: evaluation.isFormula,
      rowIndex: request.rowIndex,
      sheetId: sheet.id,
      sheetName: sheet.name,
      ...(evaluation.errorCode ? { errorCode: evaluation.errorCode } : {}),
    };
  }

  getUsedRange(sheetId?: string): UsedRangeResult {
    return getSheetUsedRange(this.#state, sheetId);
  }

  applyTransaction(request: ApplyTransactionRequest): ApplyTransactionResult {
    const execution = applyWorkbookTransaction(this.#state, request);
    const nextState =
      execution.changed && !request.dryRun
        ? {
            ...execution.state,
            hasUnsavedChanges: true,
          }
        : execution.state;
    const nextSummary = getWorkbookSummary(nextState);

    if (execution.changed && !request.dryRun) {
      this.#commitState(nextState);
    }

    return {
      changed: execution.changed,
      summary: nextSummary,
      version: nextSummary.version,
    };
  }

  async exportCsvFile(
    request: ExportCsvFileRequest,
  ): Promise<CsvFileOperationResult> {
    const filePath = normalizeCsvFilePath(request.filePath);
    const content = this.getSheetCsv(request.sheetId);

    await fs.writeFile(filePath, content, "utf8");

    const result = this.applyTransaction({
      operations: [
        {
          sheetId: request.sheetId,
          sourceFilePath: filePath,
          type: "setSheetSourceFile",
        },
      ],
    });

    return {
      ...result,
      filePath,
    };
  }

  async openWorkbookFile(
    request: OpenWorkbookFileRequest,
  ): Promise<WorkbookFileOperationResult> {
    if (this.#state.hasUnsavedChanges && !request.discardUnsavedChanges) {
      throw new Error(
        "Workbook has unsaved changes. Save it first or retry with discardUnsavedChanges: true.",
      );
    }

    const filePath = path.resolve(request.filePath);
    const content = await fs.readFile(filePath, "utf8");
    const nextState = parseWorkbookDocument(content);

    nextState.documentFilePath = filePath;
    nextState.hasUnsavedChanges = false;
    nextState.version = this.#state.version + 1;

    this.#commitState(nextState);

    const summary = getWorkbookSummary(this.#state);

    return {
      changed: true,
      filePath,
      summary,
      version: summary.version,
    };
  }

  async saveWorkbookFile(
    request: SaveWorkbookFileRequest,
  ): Promise<WorkbookFileOperationResult> {
    const filePath = normalizeWorkbookFilePath(request.filePath);

    await fs.writeFile(filePath, serializeWorkbookDocument(this.#state), "utf8");

    const result = this.#updateDocumentFilePath(filePath);

    return {
      changed: result.changed,
      filePath,
      summary: result.summary,
      version: result.summary.version,
    };
  }

  async importCsvFile(
    request: ImportCsvFileRequest,
  ): Promise<CsvFileOperationResult> {
    const filePath = path.resolve(request.filePath);
    const content = await fs.readFile(filePath, "utf8");
    const result = this.applyTransaction({
      operations: [
        {
          content,
          name: request.name,
          sheetId: request.sheetId,
          sourceFilePath: filePath,
          type: "replaceSheetFromCsv",
        },
      ],
    });

    return {
      ...result,
      filePath,
    };
  }

  #getEvaluationSnapshot(sheetId?: string): SheetEvaluationSnapshot {
    const sheet = getWorkbookSheet(this.#state, sheetId);
    const cachedSnapshot = this.#sheetEvaluationSnapshots.get(sheet.id);

    if (
      cachedSnapshot &&
      cachedSnapshot.workbookVersion === this.#state.version
    ) {
      return cachedSnapshot;
    }

    const nextSnapshot = evaluateSheet(sheet, this.#state.version);

    this.#sheetEvaluationSnapshots.set(sheet.id, nextSnapshot);
    return nextSnapshot;
  }

  #commitState(nextState: WorkbookState) {
    this.#state = nextState;
    this.#sheetEvaluationSnapshots.clear();
    this.emit("changed", getWorkbookSummary(this.#state));
  }

  #updateDocumentFilePath(filePath: string): {
    changed: boolean;
    summary: WorkbookSummary;
  } {
    if (
      this.#state.documentFilePath === filePath &&
      !this.#state.hasUnsavedChanges
    ) {
      return {
        changed: false,
        summary: getWorkbookSummary(this.#state),
      };
    }

    this.#commitState({
      ...this.#state,
      documentFilePath: filePath,
      hasUnsavedChanges: false,
      version: this.#state.version + 1,
    });

    return {
      changed: true,
      summary: getWorkbookSummary(this.#state),
    };
  }
}

function normalizeCsvFilePath(filePath: string): string {
  const resolvedFilePath = path.resolve(filePath);

  if (resolvedFilePath.toLowerCase().endsWith(".csv")) {
    return resolvedFilePath;
  }

  return `${resolvedFilePath}.csv`;
}

function normalizeWorkbookFilePath(filePath: string): string {
  const resolvedFilePath = path.resolve(filePath);

  if (resolvedFilePath.toLowerCase().endsWith(WORKBOOK_DOCUMENT_EXTENSION)) {
    return resolvedFilePath;
  }

  return `${resolvedFilePath}${WORKBOOK_DOCUMENT_EXTENSION}`;
}

function assertCellIndex(value: number, limit: number, label: string) {
  if (!Number.isInteger(value) || value < 0 || value >= limit) {
    throw new Error(
      `${label} index must be a non-negative integer within sheet bounds.`,
    );
  }
}
