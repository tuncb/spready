import { promises as fs } from 'node:fs';
import path from 'node:path';
import { EventEmitter } from 'node:events';

import {
  applyWorkbookTransaction,
  createWorkbookState,
  type CsvFileOperationResult,
  type ExportCsvFileRequest,
  getSheetCsv,
  getSheetRange,
  getSheetUsedRange,
  getWorkbookSummary,
  type ApplyTransactionRequest,
  type ApplyTransactionResult,
  type ImportCsvFileRequest,
  type SheetRangeRequest,
  type SheetRangeResult,
  type UsedRangeResult,
  type WorkbookState,
  type WorkbookSummary,
} from './workbook-core';

export class WorkbookController extends EventEmitter {
  #state: WorkbookState = createWorkbookState();

  getSummary(): WorkbookSummary {
    return getWorkbookSummary(this.#state);
  }

  getSheetCsv(sheetId?: string): string {
    return getSheetCsv(this.#state, sheetId);
  }

  getSheetRange(request: SheetRangeRequest): SheetRangeResult {
    return getSheetRange(this.#state, request);
  }

  getUsedRange(sheetId?: string): UsedRangeResult {
    return getSheetUsedRange(this.#state, sheetId);
  }

  applyTransaction(request: ApplyTransactionRequest): ApplyTransactionResult {
    const execution = applyWorkbookTransaction(this.#state, request);
    const nextState = request.dryRun ? execution.state : execution.state;
    const nextSummary = getWorkbookSummary(nextState);

    if (execution.changed && !request.dryRun) {
      this.#state = execution.state;
      this.emit('changed', nextSummary);
    }

    return {
      changed: execution.changed,
      summary: nextSummary,
      version: nextSummary.version,
    };
  }

  async exportCsvFile(request: ExportCsvFileRequest): Promise<CsvFileOperationResult> {
    const filePath = normalizeCsvFilePath(request.filePath);
    const content = this.getSheetCsv(request.sheetId);

    await fs.writeFile(filePath, content, 'utf8');

    const result = this.applyTransaction({
      operations: [
        {
          sheetId: request.sheetId,
          sourceFilePath: filePath,
          type: 'setSheetSourceFile',
        },
      ],
    });

    return {
      ...result,
      filePath,
    };
  }

  async importCsvFile(request: ImportCsvFileRequest): Promise<CsvFileOperationResult> {
    const filePath = path.resolve(request.filePath);
    const content = await fs.readFile(filePath, 'utf8');
    const result = this.applyTransaction({
      operations: [
        {
          content,
          name: request.name,
          sheetId: request.sheetId,
          sourceFilePath: filePath,
          type: 'replaceSheetFromCsv',
        },
      ],
    });

    return {
      ...result,
      filePath,
    };
  }
}

function normalizeCsvFilePath(filePath: string): string {
  const resolvedFilePath = path.resolve(filePath);

  if (resolvedFilePath.toLowerCase().endsWith('.csv')) {
    return resolvedFilePath;
  }

  return `${resolvedFilePath}.csv`;
}
