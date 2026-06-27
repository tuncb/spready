import { promises as fs } from "node:fs";
import path from "node:path";
import { EventEmitter } from "node:events";

import {
  applyWorkbookTransaction,
  buildCutRangeOperations,
  buildCreateChartOperation,
  buildFormatCellsOperations,
  buildPasteRangeOperations,
  clampSearchResultIndex,
  createClipboardRangePayload,
  createWorkbookSearchState,
  cloneWorkbookChart,
  cloneWorkbookTable,
  createWorkbookState,
  getSheetColumnCount,
  getColumnTitle,
  getWorkbookChartById,
  getWorkbookTableById,
  getWorkbookChartStatus,
  getWorkbookChartValidationIssues,
  getWorkbookSheetCharts,
  getWorkbookSheetTables,
  getWorkbookSheet,
  getSheetCsv,
  getSheetRange,
  getSheetStyleRange,
  getSheetUsedRange,
  getWorkbookSummary,
  getWorkbookRawSearchResults,
  isFormulaInput,
  getSheetRowCount,
  type ApplyTransactionRequest,
  type ApplyTransactionResult,
  type CellDataRequest,
  type CellDataResult,
  type ClearRangeRequest,
  type CopyRangeRequest,
  type CopyRangeResult,
  type CreateChartRequest,
  type CreateChartResult,
  type CreateNewWorkbookRequest,
  type CutRangeRequest,
  type CutRangeResult,
  type CsvFileOperationResult,
  type ExportCsvFileRequest,
  type FormatCellsRequest,
  type ImportCsvFileRequest,
  parseTsv,
  type OpenWorkbookFileRequest,
  type PasteRangeRequest,
  type SaveWorkbookFileRequest,
  type SetWorkbookSearchQueryRequest,
  type WorkbookChartPreview,
  type WorkbookChartResult,
  type WorkbookChartSheetReference,
  type WorkbookSheetChartPreviewsResult,
  type WorkbookSheetChartsResult,
  type SheetStyleRangeResult,
  type SheetDisplayRangeResult,
  type SheetRangeRequest,
  type SheetRangeResult,
  type UsedRangeResult,
  type WorkbookFileOperationResult,
  type WorkbookHistoryCheckoutRequest,
  type WorkbookHistoryRequest,
  type WorkbookHistoryResult,
  type WorkbookRedoRequest,
  type WorkbookConsoleOutputResult,
  type WorkbookState,
  type WorkbookSearchQuery,
  type WorkbookSearchResult,
  type WorkbookSearchState,
  type WorkbookTable,
  type WorkbookSummary,
  type WorkbookSheetTablesResult,
  type WorkbookTransactionOperation,
  type WorkbookUndoTree,
} from "./workbook-core";
import { formatWorkbookConsoleOutput } from "./workbook-console-output";
import {
  compareCellEvaluationSortValues,
  createCellKey,
  createFormulaEvaluationMetrics,
  evaluateWorkbook,
  getCellEvaluation,
  getCellKeySheetId,
  type CellEvaluation,
  type CellKey,
  type FormulaEvaluationMetrics,
  type FormulaParseCache,
  type SheetEvaluationSnapshot,
} from "./formula-engine";
import { buildWorkbookChartPreview } from "./workbook-charting";
import {
  parseWorkbookDocument,
  serializeWorkbookDocument,
  WORKBOOK_DOCUMENT_EXTENSION,
} from "./workbook-document";

interface WorkbookHistoryNode {
  childIds: string[];
  id: string;
  parentId?: string;
  state: WorkbookState;
}

type EvaluationTraceSink = (message: string) => void;

export interface WorkbookEvaluationMetricsEvent {
  durationMs: number;
  metrics: FormulaEvaluationMetrics;
  reason: string;
  requestedSheetId: string;
  snapshotCount: number;
  volatileSnapshotCount: number;
  workbookVersion: number;
}

type EvaluationMetricsSink = (event: WorkbookEvaluationMetricsEvent) => void;

export class WorkbookController extends EventEmitter {
  #state: WorkbookState = createWorkbookState();
  #sheetEvaluationSnapshots = new Map<string, SheetEvaluationSnapshot>();
  #formulaParseCache: FormulaParseCache = new Map();
  #evaluationClock: () => Date;
  #evaluationMetricsSink?: EvaluationMetricsSink;
  #evaluationTraceSink?: EvaluationTraceSink;
  #performanceClock: () => number;
  #historyNodes = new Map<string, WorkbookHistoryNode>();
  #currentHistoryNodeId = "";
  #rootHistoryNodeId = "";
  #savedHistoryNodeId: string | undefined;
  #nextHistoryNodeNumber = 1;
  #searchQuery: WorkbookSearchQuery = {
    activeResultIndex: -1,
    scope: "sheet",
    text: "",
    valueMode: "display",
  };

  constructor(
    options: {
      evaluationClock?: () => Date;
      evaluationMetricsSink?: EvaluationMetricsSink;
      evaluationTraceSink?: EvaluationTraceSink;
      performanceClock?: () => number;
    } = {},
  ) {
    super();
    this.#evaluationClock = options.evaluationClock ?? (() => new Date());
    this.#evaluationMetricsSink = options.evaluationMetricsSink;
    this.#evaluationTraceSink =
      options.evaluationTraceSink ?? createEnvironmentEvaluationTraceSink();
    this.#performanceClock = options.performanceClock ?? Date.now;
    this.#resetHistory(this.#state, true);
  }

  getSummary(): WorkbookSummary {
    return getWorkbookSummary(this.#state);
  }

  getUndoTree(): WorkbookUndoTree {
    return this.#buildUndoTree();
  }

  getSearchState(): WorkbookSearchState {
    return this.#buildSearchState();
  }

  setSearchQuery(request: SetWorkbookSearchQueryRequest): WorkbookSearchState {
    this.#searchQuery = {
      activeResultIndex: request.text.length > 0 ? 0 : -1,
      scope: request.scope,
      text: request.text,
      valueMode: request.valueMode ?? this.#searchQuery.valueMode,
    };

    return this.#buildSearchState();
  }

  clearSearch(): WorkbookSearchState {
    this.#searchQuery = {
      activeResultIndex: -1,
      scope: this.#searchQuery.scope,
      text: "",
      valueMode: this.#searchQuery.valueMode,
    };

    return this.#buildSearchState();
  }

  setActiveSearchResult(index: number): WorkbookSearchState {
    const state = this.#buildSearchState();

    this.#searchQuery = {
      ...state.query,
      activeResultIndex: clampSearchResultIndex(index, state.results.length),
    };

    return this.#buildSearchState();
  }

  goToNextSearchResult(): WorkbookSearchState {
    return this.#moveSearchResult(1);
  }

  goToPreviousSearchResult(): WorkbookSearchState {
    return this.#moveSearchResult(-1);
  }

  undo(request: WorkbookHistoryRequest = {}): WorkbookHistoryResult {
    this.#assertExpectedVersion(request.expectedVersion);

    const currentNode = this.#getCurrentHistoryNode();

    if (!currentNode.parentId) {
      throw new Error("No undo history is available.");
    }

    return this.#restoreHistoryNode(currentNode.parentId);
  }

  redo(request: WorkbookRedoRequest = {}): WorkbookHistoryResult {
    this.#assertExpectedVersion(request.expectedVersion);

    const currentNode = this.#getCurrentHistoryNode();
    const targetNodeId = request.nodeId ?? currentNode.childIds[currentNode.childIds.length - 1];

    if (!targetNodeId) {
      throw new Error("No redo history is available.");
    }

    if (!currentNode.childIds.includes(targetNodeId)) {
      throw new Error(`History node "${targetNodeId}" is not a redo child of the current node.`);
    }

    return this.#restoreHistoryNode(targetNodeId);
  }

  checkoutUndoNode(request: WorkbookHistoryCheckoutRequest): WorkbookHistoryResult {
    this.#assertExpectedVersion(request.expectedVersion);

    if (!this.#historyNodes.has(request.nodeId)) {
      throw new Error(`History node "${request.nodeId}" was not found.`);
    }

    return this.#restoreHistoryNode(request.nodeId);
  }

  getSheetCharts(sheetId?: string): WorkbookSheetChartsResult {
    return getWorkbookSheetCharts(this.#state, sheetId);
  }

  getSheetChartPreviews(sheetId?: string): WorkbookSheetChartPreviewsResult {
    const sheetCharts = getWorkbookSheetCharts(this.#state, sheetId);

    return {
      previews: sheetCharts.charts.map((chart) => {
        const sourceSheet = this.#state.sheets.find(
          (sheet) => sheet.id === chart.spec.source.range.sheetId,
        );

        return buildWorkbookChartPreview(
          chart,
          sourceSheet,
          sourceSheet
            ? this.#getEvaluationSnapshot(sourceSheet.id, "getSheetChartPreviews")
            : undefined,
          this.#getChartSheetReferences(),
        );
      }),
      sheetId: sheetCharts.sheetId,
      sheetName: sheetCharts.sheetName,
    };
  }

  getChart(chartId: string): WorkbookChartResult {
    const chart = getWorkbookChartById(this.#state, chartId);

    return {
      chart: cloneWorkbookChart(chart),
      status: getWorkbookChartStatus(chart, this.#getChartSheetReferences()),
      validationIssues: getWorkbookChartValidationIssues(chart, this.#getChartSheetReferences()),
    };
  }

  getChartPreview(chartId: string): WorkbookChartPreview {
    const chart = getWorkbookChartById(this.#state, chartId);
    const sourceSheet = this.#state.sheets.find(
      (sheet) => sheet.id === chart.spec.source.range.sheetId,
    );

    return buildWorkbookChartPreview(
      cloneWorkbookChart(chart),
      sourceSheet,
      sourceSheet ? this.#getEvaluationSnapshot(sourceSheet.id, "getChartPreview") : undefined,
      this.#getChartSheetReferences(),
    );
  }

  getSheetTables(sheetId?: string): WorkbookSheetTablesResult {
    return getWorkbookSheetTables(this.#state, sheetId);
  }

  getTable(tableId: string): WorkbookTable {
    return cloneWorkbookTable(getWorkbookTableById(this.#state, tableId));
  }

  getSheetCsv(sheetId?: string): string {
    return getSheetCsv(this.#state, sheetId);
  }

  getSheetRange(request: SheetRangeRequest): SheetRangeResult {
    return getSheetRange(this.#state, request);
  }

  getSheetStyleRange(request: SheetRangeRequest): SheetStyleRangeResult {
    return getSheetStyleRange(this.#state, request);
  }

  getSheetDisplayRange(request: SheetRangeRequest): SheetDisplayRangeResult {
    const rawRange = getSheetRange(this.#state, request);
    const snapshot = this.#getEvaluationSnapshot(rawRange.sheetId, "getSheetDisplayRange");

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
      this.#getEvaluationSnapshot(sheet.id, "getCellData"),
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
      ...(sheet.cellStyles[`${request.rowIndex}:${request.columnIndex}`]
        ? {
            style: {
              ...sheet.cellStyles[`${request.rowIndex}:${request.columnIndex}`],
            },
          }
        : {}),
    };
  }

  getUsedRange(sheetId?: string): UsedRangeResult {
    return getSheetUsedRange(this.#state, sheetId);
  }

  getConsoleOutput(): WorkbookConsoleOutputResult {
    const summary = this.getSummary();
    const sheets = summary.sheets.map((sheet) => {
      const usedRange = this.getUsedRange(sheet.id);

      return {
        ...(usedRange.rowCount > 0 && usedRange.columnCount > 0
          ? {
              displayRange: this.getSheetDisplayRange({
                columnCount: usedRange.columnCount,
                rowCount: usedRange.rowCount,
                sheetId: sheet.id,
                startColumn: usedRange.startColumn,
                startRow: usedRange.startRow,
              }),
            }
          : {}),
        sheetId: sheet.id,
        sheetName: sheet.name,
        usedRange,
      };
    });

    return {
      text: formatWorkbookConsoleOutput(summary, sheets),
    };
  }

  copyRange(request: CopyRangeRequest): CopyRangeResult {
    const mode = request.mode ?? "raw";
    const rawRange = this.getSheetRange(request);
    const displayRange = this.getSheetDisplayRange(request);
    const clipboard = createClipboardRangePayload(this.#state, rawRange, displayRange);
    const range = mode === "display" ? displayRange : rawRange;

    return {
      ...range,
      clipboard,
      mode,
      text: mode === "display" ? clipboard.displayText : clipboard.rawText,
    };
  }

  cutRange(request: CutRangeRequest): CutRangeResult {
    const mode = request.mode ?? "raw";
    const rawRange = this.getSheetRange(request);
    const displayRange = this.getSheetDisplayRange(request);
    const clipboard = createClipboardRangePayload(this.#state, rawRange, displayRange);
    const selectedRange = mode === "display" ? displayRange : rawRange;
    const clearResult = this.applyTransaction({
      operations: buildCutRangeOperations(this.#state, rawRange),
    });

    return {
      ...selectedRange,
      changed: clearResult.changed,
      clipboard,
      mode,
      summary: clearResult.summary,
      text: mode === "display" ? clipboard.displayText : clipboard.rawText,
      version: clearResult.version,
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

  createChart(request: CreateChartRequest): CreateChartResult {
    const { chartId, operation } = buildCreateChartOperation(this.#state, request);
    const result = this.applyTransaction({
      dryRun: request.dryRun,
      expectedVersion: request.expectedVersion,
      operations: [operation],
    });
    const chart = result.summary.charts.find((entry) => entry.id === chartId);

    if (!chart) {
      throw new Error(`Chart "${chartId}" was not found after creation.`);
    }

    return {
      ...result,
      chart,
    };
  }

  formatCells(request: FormatCellsRequest): ApplyTransactionResult {
    return this.applyTransaction({
      dryRun: request.dryRun,
      expectedVersion: request.expectedVersion,
      operations: buildFormatCellsOperations(this.#state, request),
    });
  }

  pasteRange(request: PasteRangeRequest): ApplyTransactionResult {
    const values =
      request.values?.map((row) => [...row]) ??
      (request.clipboard
        ? request.mode === "display"
          ? request.clipboard.displayValues.map((row) => [...row])
          : request.clipboard.rawValues.map((row) => [...row])
        : undefined) ??
      (request.text !== undefined ? parseTsv(request.text) : undefined);

    if (!values || values.length === 0) {
      return this.applyTransaction({
        operations: [],
      });
    }

    return this.applyTransaction({
      operations: buildPasteRangeOperations(this.#state, request, values),
    });
  }

  createNewWorkbook(request: CreateNewWorkbookRequest = {}): ApplyTransactionResult {
    if (this.#state.hasUnsavedChanges && !request.discardUnsavedChanges) {
      throw new Error(
        "Workbook has unsaved changes. Save it first or retry with discardUnsavedChanges: true.",
      );
    }

    const nextState = createWorkbookState();

    nextState.version = this.#state.version + 1;
    this.#resetHistory(nextState, true);
    this.#commitState(nextState);

    const summary = getWorkbookSummary(this.#state);

    return {
      changed: true,
      summary,
      version: summary.version,
    };
  }

  applyTransaction(request: ApplyTransactionRequest): ApplyTransactionResult {
    this.#assertExpectedVersion(request.expectedVersion);

    const preparedRequest = this.#prepareTransaction(request);
    const execution = applyWorkbookTransaction(this.#state, preparedRequest);
    const nextState =
      execution.changed && !request.dryRun
        ? {
            ...execution.state,
            hasUnsavedChanges: true,
          }
        : execution.state;
    const nextSummary = getWorkbookSummary(nextState);

    if (execution.changed && !request.dryRun) {
      const nextEvaluationSnapshots = this.#getTransactionEvaluationSnapshots(
        preparedRequest.operations,
      );

      this.#recordHistoryNode(nextState);
      this.#commitState(nextState, {
        evaluationSnapshots: nextEvaluationSnapshots,
        preserveEvaluationSnapshots: shouldPreserveEvaluationSnapshots(preparedRequest.operations),
      });
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

  async openWorkbookFile(request: OpenWorkbookFileRequest): Promise<WorkbookFileOperationResult> {
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

    this.#resetHistory(nextState, true);
    this.#commitState(nextState);

    const summary = getWorkbookSummary(this.#state);

    return {
      changed: true,
      filePath,
      summary,
      version: summary.version,
    };
  }

  async saveWorkbookFile(request: SaveWorkbookFileRequest): Promise<WorkbookFileOperationResult> {
    const filePath = normalizeWorkbookFilePath(request.filePath);

    await fs.writeFile(filePath, serializeWorkbookDocument(this.#state), "utf8");

    const result = this.#markCurrentHistoryNodeSaved(filePath);

    return {
      changed: result.changed,
      filePath,
      summary: result.summary,
      version: result.summary.version,
    };
  }

  async importCsvFile(request: ImportCsvFileRequest): Promise<CsvFileOperationResult> {
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

  #prepareTransaction(request: ApplyTransactionRequest): ApplyTransactionRequest {
    let operations: WorkbookTransactionOperation[] | undefined;

    request.operations.forEach((operation, index) => {
      if (
        operation.type !== "sortTable" ||
        (operation.valueMode ?? "raw") !== "display" ||
        operation.bodyRowOrder !== undefined
      ) {
        return;
      }

      operations ??= [...request.operations];
      operations[index] = {
        ...operation,
        bodyRowOrder: this.#getDisplayTableSortBodyRowOrder(operation),
      };
    });

    return operations
      ? {
          ...request,
          operations,
        }
      : request;
  }

  #moveSearchResult(direction: 1 | -1): WorkbookSearchState {
    const state = this.#buildSearchState();

    if (state.results.length === 0) {
      this.#searchQuery = state.query;
      return state;
    }

    const currentIndex = state.query.activeResultIndex >= 0 ? state.query.activeResultIndex : 0;
    const activeResultIndex =
      (currentIndex + direction + state.results.length) % state.results.length;

    this.#searchQuery = {
      ...state.query,
      activeResultIndex,
    };

    return this.#buildSearchState();
  }

  #buildSearchState(): WorkbookSearchState {
    const state = createWorkbookSearchState(
      this.#searchQuery,
      this.#getSearchResults(this.#searchQuery),
    );

    this.#searchQuery = state.query;
    return state;
  }

  #getSearchResults(query: WorkbookSearchQuery): WorkbookSearchResult[] {
    if (query.valueMode === "raw") {
      return getWorkbookRawSearchResults(this.#state, query.text, query.scope);
    }

    const needle = query.text.toLocaleLowerCase();

    if (needle.length === 0) {
      return [];
    }

    const sheets =
      query.scope === "workbook"
        ? this.#state.sheets
        : this.#state.sheets.filter((sheet) => sheet.id === this.#state.activeSheetId);
    const results: WorkbookSearchResult[] = [];

    for (const sheet of sheets) {
      const snapshot = this.#getEvaluationSnapshot(sheet.id, "getSearchState");

      for (let rowIndex = 0; rowIndex < getSheetRowCount(sheet); rowIndex += 1) {
        for (let columnIndex = 0; columnIndex < getSheetColumnCount(sheet); columnIndex += 1) {
          const displayValue = getCellEvaluation(snapshot, rowIndex, columnIndex).display;

          if (displayValue.length === 0 || !displayValue.toLocaleLowerCase().includes(needle)) {
            continue;
          }

          results.push({
            address: `${getColumnTitle(columnIndex)}${rowIndex + 1}`,
            columnIndex,
            matchedText: displayValue,
            rowIndex,
            sheetId: sheet.id,
            sheetName: sheet.name,
          });
        }
      }
    }

    return results;
  }

  #getDisplayTableSortBodyRowOrder(
    operation: Extract<WorkbookTransactionOperation, { type: "sortTable" }>,
  ): number[] {
    const table = getWorkbookTableById(this.#state, operation.tableId);
    const headerOffset = table.hasHeaderRow ? 1 : 0;
    const bodyStartRow = table.range.startRow + headerOffset;
    const bodyRowCount = table.range.rowCount - headerOffset;

    if (bodyRowCount <= 1) {
      return Array.from(
        { length: Math.max(0, bodyRowCount) },
        (_value, index) => bodyStartRow + index,
      );
    }

    const snapshot = this.#getEvaluationSnapshot(table.range.sheetId, "sortTable");
    const rows = Array.from({ length: bodyRowCount }, (_value, index) => ({
      originalIndex: index,
      rowIndex: bodyStartRow + index,
    }));

    return rows
      .sort((left, right) => {
        for (const key of operation.keys) {
          const comparison = compareDisplayTableSortValues(
            getCellEvaluation(snapshot, left.rowIndex, key.columnIndex),
            getCellEvaluation(snapshot, right.rowIndex, key.columnIndex),
            key.direction,
          );

          if (comparison !== 0) {
            return comparison;
          }
        }

        return left.originalIndex - right.originalIndex;
      })
      .map((entry) => entry.rowIndex);
  }

  #getEvaluationSnapshot(
    sheetId?: string,
    reason = "getEvaluationSnapshot",
  ): SheetEvaluationSnapshot {
    const sheet = getWorkbookSheet(this.#state, sheetId);
    const cachedSnapshot = this.#sheetEvaluationSnapshots.get(sheet.id);

    if (
      cachedSnapshot &&
      cachedSnapshot.workbookVersion === this.#state.version &&
      !cachedSnapshot.hasVolatileFunctions
    ) {
      return cachedSnapshot;
    }

    const metrics =
      this.#evaluationTraceSink || this.#evaluationMetricsSink
        ? createFormulaEvaluationMetrics()
        : undefined;
    const startedMs = this.#performanceClock();
    const nextSnapshots = evaluateWorkbook(this.#state, this.#state.version, {
      functionClock: metrics ? this.#performanceClock : undefined,
      metrics,
      now: this.#evaluationClock(),
      parseCache: this.#formulaParseCache,
      seedSnapshots: this.#canSeedEvaluationSnapshots()
        ? this.#sheetEvaluationSnapshots
        : undefined,
    });
    const durationMs = this.#performanceClock() - startedMs;

    this.#sheetEvaluationSnapshots = nextSnapshots;
    if (metrics) {
      this.#traceEvaluation(reason, sheet.id, metrics, durationMs);
      this.#emitEvaluationMetrics(reason, sheet.id, metrics, durationMs);
    }
    const nextSnapshot = this.#sheetEvaluationSnapshots.get(sheet.id);

    if (!nextSnapshot) {
      throw new Error(`Sheet "${sheet.id}" was not found in the evaluation snapshot.`);
    }

    return nextSnapshot;
  }

  #traceEvaluation(
    reason: string,
    requestedSheetId: string,
    metrics: FormulaEvaluationMetrics,
    durationMs: number,
  ) {
    if (!this.#evaluationTraceSink) {
      return;
    }

    const snapshotCount = this.#sheetEvaluationSnapshots.size;
    const volatileSnapshotCount = [...this.#sheetEvaluationSnapshots.values()].filter(
      (snapshot) => snapshot.hasVolatileFunctions,
    ).length;

    this.#evaluationTraceSink(
      formatEvaluationTraceLog(reason, requestedSheetId, this.#state.version, durationMs, metrics, {
        snapshotCount,
        volatileSnapshotCount,
      }),
    );
  }

  #emitEvaluationMetrics(
    reason: string,
    requestedSheetId: string,
    metrics: FormulaEvaluationMetrics,
    durationMs: number,
  ) {
    if (!this.#evaluationMetricsSink) {
      return;
    }

    const snapshotCount = this.#sheetEvaluationSnapshots.size;
    const volatileSnapshotCount = [...this.#sheetEvaluationSnapshots.values()].filter(
      (snapshot) => snapshot.hasVolatileFunctions,
    ).length;

    this.#evaluationMetricsSink({
      durationMs,
      metrics,
      reason,
      requestedSheetId,
      snapshotCount,
      volatileSnapshotCount,
      workbookVersion: this.#state.version,
    });
  }

  #canSeedEvaluationSnapshots() {
    return ![...this.#sheetEvaluationSnapshots.values()].some(
      (snapshot) => snapshot.hasVolatileFunctions,
    );
  }

  #getChartSheetReferences(): WorkbookChartSheetReference[] {
    return this.#state.sheets.map((sheet) => ({
      columnCount: getSheetColumnCount(sheet),
      id: sheet.id,
      rowCount: getSheetRowCount(sheet),
    }));
  }

  #getTransactionEvaluationSnapshots(
    operations: readonly WorkbookTransactionOperation[],
  ): Map<string, SheetEvaluationSnapshot> | undefined {
    const changedCellKeys = getIncrementalCellWriteKeys(this.#state, operations);

    if (!changedCellKeys || this.#sheetEvaluationSnapshots.size === 0) {
      return undefined;
    }

    if (
      [...this.#sheetEvaluationSnapshots.values()].some((snapshot) => snapshot.hasVolatileFunctions)
    ) {
      return undefined;
    }

    return invalidateEvaluationSnapshots(this.#sheetEvaluationSnapshots, changedCellKeys);
  }

  #commitState(
    nextState: WorkbookState,
    options: {
      evaluationSnapshots?: Map<string, SheetEvaluationSnapshot>;
      preserveEvaluationSnapshots?: boolean;
    } = {},
  ) {
    this.#state = nextState;

    if (options.evaluationSnapshots) {
      this.#sheetEvaluationSnapshots = options.evaluationSnapshots;
    } else if (options.preserveEvaluationSnapshots) {
      this.#sheetEvaluationSnapshots = retagEvaluationSnapshots(
        this.#sheetEvaluationSnapshots,
        nextState.version,
      );
    } else {
      this.#sheetEvaluationSnapshots.clear();
    }

    this.emit("changed", getWorkbookSummary(this.#state));
  }

  #assertExpectedVersion(expectedVersion: number | undefined) {
    if (expectedVersion !== undefined && expectedVersion !== this.#state.version) {
      throw new Error(
        `Expected workbook version ${expectedVersion}, but current version is ${this.#state.version}.`,
      );
    }
  }

  #buildUndoTree(): WorkbookUndoTree {
    const currentNode = this.#getCurrentHistoryNode();

    return {
      canRedo: currentNode.childIds.length > 0,
      canUndo: currentNode.parentId !== undefined,
      currentNodeId: this.#currentHistoryNodeId,
      nodes: [...this.#historyNodes.values()].map((node) => ({
        childIds: [...node.childIds],
        id: node.id,
        isCurrent: node.id === this.#currentHistoryNodeId,
        isSaved: node.id === this.#savedHistoryNodeId,
        parentId: node.parentId,
        summary: this.#getHistoryNodeSummary(node),
      })),
      rootNodeId: this.#rootHistoryNodeId,
      savedNodeId: this.#savedHistoryNodeId,
    };
  }

  #getHistoryNodeSummary(node: WorkbookHistoryNode): WorkbookSummary {
    return getWorkbookSummary({
      ...node.state,
      documentFilePath: this.#state.documentFilePath,
      hasUnsavedChanges: node.id !== this.#savedHistoryNodeId,
    });
  }

  #getCurrentHistoryNode(): WorkbookHistoryNode {
    const node = this.#historyNodes.get(this.#currentHistoryNodeId);

    if (!node) {
      throw new Error("Workbook history is not initialized.");
    }

    return node;
  }

  #recordHistoryNode(state: WorkbookState) {
    const currentNode = this.#getCurrentHistoryNode();
    const nodeId = this.#createHistoryNodeId();
    const node: WorkbookHistoryNode = {
      childIds: [],
      id: nodeId,
      parentId: currentNode.id,
      state: {
        ...state,
        hasUnsavedChanges: true,
      },
    };

    currentNode.childIds.push(node.id);
    this.#historyNodes.set(node.id, node);
    this.#currentHistoryNodeId = node.id;
  }

  #resetHistory(state: WorkbookState, isSaved: boolean) {
    const nodeId = this.#createHistoryNodeId();
    const node: WorkbookHistoryNode = {
      childIds: [],
      id: nodeId,
      state: {
        ...state,
        hasUnsavedChanges: !isSaved,
      },
    };

    this.#historyNodes = new Map([[node.id, node]]);
    this.#rootHistoryNodeId = node.id;
    this.#currentHistoryNodeId = node.id;
    this.#savedHistoryNodeId = isSaved ? node.id : undefined;
    this.#state = node.state;
  }

  #restoreHistoryNode(nodeId: string): WorkbookHistoryResult {
    const targetNode = this.#historyNodes.get(nodeId);

    if (!targetNode) {
      throw new Error(`History node "${nodeId}" was not found.`);
    }

    if (nodeId === this.#currentHistoryNodeId) {
      return {
        changed: false,
        summary: getWorkbookSummary(this.#state),
        undoTree: this.#buildUndoTree(),
        version: this.#state.version,
      };
    }

    const nextState: WorkbookState = {
      ...targetNode.state,
      documentFilePath: this.#state.documentFilePath,
      hasUnsavedChanges: nodeId !== this.#savedHistoryNodeId,
      version: this.#state.version + 1,
    };

    targetNode.state = nextState;
    this.#currentHistoryNodeId = nodeId;
    this.#commitState(nextState);

    const summary = getWorkbookSummary(this.#state);

    return {
      changed: true,
      summary,
      undoTree: this.#buildUndoTree(),
      version: summary.version,
    };
  }

  #markCurrentHistoryNodeSaved(filePath: string): {
    changed: boolean;
    summary: WorkbookSummary;
  } {
    const currentNode = this.#getCurrentHistoryNode();

    if (
      this.#state.documentFilePath === filePath &&
      this.#savedHistoryNodeId === currentNode.id &&
      !this.#state.hasUnsavedChanges
    ) {
      return {
        changed: false,
        summary: getWorkbookSummary(this.#state),
      };
    }

    this.#savedHistoryNodeId = currentNode.id;

    const nextState: WorkbookState = {
      ...this.#state,
      documentFilePath: filePath,
      hasUnsavedChanges: false,
      version: this.#state.version + 1,
    };

    currentNode.state = nextState;
    this.#commitState(nextState);

    return {
      changed: true,
      summary: getWorkbookSummary(this.#state),
    };
  }

  #createHistoryNodeId(): string {
    const nodeId = `history-${this.#nextHistoryNodeNumber}`;

    this.#nextHistoryNodeNumber += 1;
    return nodeId;
  }
}

function shouldPreserveEvaluationSnapshots(operations: readonly WorkbookTransactionOperation[]) {
  return (
    operations.length > 0 && operations.every((operation) => operation.type === "setActiveSheet")
  );
}

function getIncrementalCellWriteKeys(
  state: WorkbookState,
  operations: readonly WorkbookTransactionOperation[],
): Set<CellKey> | undefined {
  const changedCellKeys = new Set<CellKey>();

  for (const operation of operations) {
    if (operation.type === "setActiveSheet") {
      continue;
    }

    if (operation.type !== "setCell") {
      return undefined;
    }

    const sheet = getWorkbookSheet(state, operation.sheetId);

    if (
      operation.rowIndex < 0 ||
      operation.columnIndex < 0 ||
      operation.rowIndex >= getSheetRowCount(sheet) ||
      operation.columnIndex >= getSheetColumnCount(sheet)
    ) {
      return undefined;
    }

    const previousInput = sheet.cells[operation.rowIndex]?.[operation.columnIndex] ?? "";

    if (isFormulaInput(previousInput) || isFormulaInput(operation.value)) {
      return undefined;
    }

    changedCellKeys.add(createCellKey(sheet.id, operation.rowIndex, operation.columnIndex));
  }

  return changedCellKeys.size > 0 ? changedCellKeys : undefined;
}

function invalidateEvaluationSnapshots(
  snapshots: Map<string, SheetEvaluationSnapshot>,
  changedCellKeys: ReadonlySet<CellKey>,
): Map<string, SheetEvaluationSnapshot> {
  const dirtyCellKeys = getTransitiveDirtyCellKeys(snapshots, changedCellKeys);
  const nextSnapshots = cloneEvaluationSnapshots(snapshots);

  for (const dirtyCellKey of dirtyCellKeys) {
    const dirtySnapshot = nextSnapshots.get(getCellKeySheetId(dirtyCellKey));

    if (!dirtySnapshot) {
      continue;
    }

    dirtySnapshot.cells.delete(dirtyCellKey);
  }

  return nextSnapshots;
}

function getTransitiveDirtyCellKeys(
  snapshots: Map<string, SheetEvaluationSnapshot>,
  changedCellKeys: ReadonlySet<CellKey>,
): Set<CellKey> {
  const dirtyCellKeys = new Set(changedCellKeys);
  const queue = [...changedCellKeys];

  for (let index = 0; index < queue.length; index += 1) {
    const cellKey = queue[index];
    const snapshot = snapshots.get(getCellKeySheetId(cellKey));
    const dependents = snapshot?.dependents.get(cellKey);

    if (!dependents) {
      continue;
    }

    for (const dependentKey of dependents) {
      if (dirtyCellKeys.has(dependentKey)) {
        continue;
      }

      dirtyCellKeys.add(dependentKey);
      queue.push(dependentKey);
    }
  }

  return dirtyCellKeys;
}

function cloneEvaluationSnapshots(
  snapshots: Map<string, SheetEvaluationSnapshot>,
): Map<string, SheetEvaluationSnapshot> {
  return new Map(
    [...snapshots].map(([sheetId, snapshot]) => [
      sheetId,
      {
        ...snapshot,
        cells: new Map(snapshot.cells),
        dependents: cloneCellKeySetMap(snapshot.dependents),
        precedents: cloneCellKeySetMap(snapshot.precedents),
      },
    ]),
  );
}

function cloneCellKeySetMap(map: Map<CellKey, Set<CellKey>>): Map<CellKey, Set<CellKey>> {
  return new Map([...map].map(([cellKey, values]) => [cellKey, new Set(values)]));
}

function createEnvironmentEvaluationTraceSink(): EvaluationTraceSink | undefined {
  return process.env.SPREADY_EVALUATION_TRACE === "1"
    ? (message) => console.error(message)
    : undefined;
}

function formatEvaluationTraceLog(
  reason: string,
  requestedSheetId: string,
  workbookVersion: number,
  durationMs: number,
  metrics: FormulaEvaluationMetrics,
  snapshotSummary: { snapshotCount: number; volatileSnapshotCount: number },
) {
  return [
    "[spready-evaluation]",
    `reason=${reason}`,
    `version=${workbookVersion}`,
    `requestedSheetId=${requestedSheetId}`,
    `durationMs=${Math.round(durationMs)}`,
    `cellsEvaluated=${metrics.cellsEvaluated}`,
    `formulasParsed=${metrics.formulasParsed}`,
    `rangeCellsMaterialized=${metrics.rangeCellsMaterialized}`,
    `dependencyKeysRecorded=${metrics.dependencyKeysRecorded}`,
    `snapshots=${snapshotSummary.snapshotCount}`,
    `volatileSnapshots=${snapshotSummary.volatileSnapshotCount}`,
  ].join(" ");
}

function retagEvaluationSnapshots(
  snapshots: Map<string, SheetEvaluationSnapshot>,
  workbookVersion: number,
) {
  return new Map(
    [...snapshots].map(([sheetId, snapshot]) => [
      sheetId,
      {
        ...snapshot,
        workbookVersion,
      },
    ]),
  );
}

function compareDisplayTableSortValues(
  left: CellEvaluation,
  right: CellEvaluation,
  direction: "ascending" | "descending",
): number {
  const leftBlank = left.value.type === "blank";
  const rightBlank = right.value.type === "blank";

  if (leftBlank || rightBlank) {
    return compareCellEvaluationSortValues(left, right);
  }

  const comparison = compareCellEvaluationSortValues(left, right);
  return direction === "ascending" ? comparison : -comparison;
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
    throw new Error(`${label} index must be a non-negative integer within sheet bounds.`);
  }
}
