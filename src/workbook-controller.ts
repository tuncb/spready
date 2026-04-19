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
  parseTsv,
  serializeTsv,
  type ClearRangeRequest,
  type SheetRangeRequest,
  type SheetDisplayRangeResult,
  type SheetRangeResult,
  type CopyRangeRequest,
  type CopyRangeResult,
  type PasteRangeRequest,
  type UsedRangeResult,
  type WorkbookState,
  type WorkbookSummary,
} from "./workbook-core";
import {
  evaluateSheet,
  getCellEvaluation,
  type SheetEvaluationSnapshot,
} from "./formula-engine";

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

  copyRange(request: CopyRangeRequest): CopyRangeResult {
    const mode = request.mode ?? "raw";
    const range =
      mode === "display"
        ? this.getSheetDisplayRange(request)
        : this.getSheetRange(request);

    return {
      ...range,
      mode,
      text: serializeTsv(range.values),
    };
  }

  clearRange(request: ClearRangeRequest): ApplyTransactionResult {
    return this.applyTransaction({
      operations: [
        {
          ...request,
          type: "clearRange",
        },
      ],
    });
  }

  pasteRange(request: PasteRangeRequest): ApplyTransactionResult {
    const values =
      request.values?.map((row) => [...row]) ??
      (request.text !== undefined ? parseTsv(request.text) : undefined);

    if (!values || values.length === 0) {
      return this.applyTransaction({
        operations: [],
      });
    }

    return this.applyTransaction({
      operations: [
        {
          sheetId: request.sheetId,
          startColumn: request.startColumn,
          startRow: request.startRow,
          type: "setRange",
          values,
        },
      ],
    });
  }

  applyTransaction(request: ApplyTransactionRequest): ApplyTransactionResult {
    const execution = applyWorkbookTransaction(this.#state, request);
    const nextSummary = getWorkbookSummary(execution.state);

    if (execution.changed && !request.dryRun) {
      this.#state = execution.state;
      this.#sheetEvaluationSnapshots.clear();
      this.emit("changed", nextSummary);
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
}

function normalizeCsvFilePath(filePath: string): string {
  const resolvedFilePath = path.resolve(filePath);

  if (resolvedFilePath.toLowerCase().endsWith(".csv")) {
    return resolvedFilePath;
  }

  return `${resolvedFilePath}.csv`;
}

function assertCellIndex(value: number, limit: number, label: string) {
  if (!Number.isInteger(value) || value < 0 || value >= limit) {
    throw new Error(
      `${label} index must be a non-negative integer within sheet bounds.`,
    );
  }
}
