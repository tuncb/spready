import { EventEmitter } from 'node:events';

import {
  applyWorkbookTransaction,
  createWorkbookState,
  getSheetCsv,
  getSheetRange,
  getSheetUsedRange,
  getWorkbookSummary,
  type ApplyTransactionRequest,
  type ApplyTransactionResult,
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
}
