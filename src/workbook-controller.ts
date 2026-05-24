import { promises as fs } from "node:fs";
import path from "node:path";
import { EventEmitter } from "node:events";

import {
  applyWorkbookTransaction,
  buildCreateChartOperation,
  buildFormatCellsOperations,
  cloneWorkbookChart,
  createWorkbookState,
  getSheetColumnCount,
  getWorkbookChartById,
  getWorkbookChartStatus,
  getWorkbookChartValidationIssues,
  getWorkbookSheetCharts,
  getWorkbookSheet,
  getSheetCsv,
  getSheetRange,
  getSheetStyleRange,
  getSheetUsedRange,
  getWorkbookSummary,
  getSheetRowCount,
  type ApplyTransactionRequest,
  type ApplyTransactionResult,
  type CellDataRequest,
  type CellDataResult,
  type ClipboardRangePayload,
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
  serializeTsv,
  type OpenWorkbookFileRequest,
  type PasteRangeRequest,
  type SaveWorkbookFileRequest,
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
  type WorkbookState,
  type WorkbookSummary,
  type WorkbookUndoTree,
} from "./workbook-core";
import {
  evaluateWorkbook,
  getCellEvaluation,
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

export class WorkbookController extends EventEmitter {
  #state: WorkbookState = createWorkbookState();
  #sheetEvaluationSnapshots = new Map<string, SheetEvaluationSnapshot>();
  #evaluationClock: () => Date;
  #historyNodes = new Map<string, WorkbookHistoryNode>();
  #currentHistoryNodeId = "";
  #rootHistoryNodeId = "";
  #savedHistoryNodeId: string | undefined;
  #nextHistoryNodeNumber = 1;

  constructor(options: { evaluationClock?: () => Date } = {}) {
    super();
    this.#evaluationClock = options.evaluationClock ?? (() => new Date());
    this.#resetHistory(this.#state, true);
  }

  getSummary(): WorkbookSummary {
    return getWorkbookSummary(this.#state);
  }

  getUndoTree(): WorkbookUndoTree {
    return this.#buildUndoTree();
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
          sourceSheet ? this.#getEvaluationSnapshot(sourceSheet.id) : undefined,
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
      sourceSheet ? this.#getEvaluationSnapshot(sourceSheet.id) : undefined,
      this.#getChartSheetReferences(),
    );
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

  copyRange(request: CopyRangeRequest): CopyRangeResult {
    const mode = request.mode ?? "raw";
    const range =
      mode === "display" ? this.getSheetDisplayRange(request) : this.getSheetRange(request);

    return {
      ...range,
      mode,
      text: serializeTsv(range.values),
    };
  }

  cutRange(request: CutRangeRequest): CutRangeResult {
    const mode = request.mode ?? "raw";
    const rawRange = this.getSheetRange(request);
    const displayRange = this.getSheetDisplayRange(request);
    const clipboard: ClipboardRangePayload = {
      displayText: serializeTsv(displayRange.values),
      displayValues: cloneRangeValues(displayRange.values),
      rawText: serializeTsv(rawRange.values),
      rawValues: cloneRangeValues(rawRange.values),
    };
    const selectedRange = mode === "display" ? displayRange : rawRange;
    const clearResult = this.clearRange({
      columnCount: rawRange.columnCount,
      rowCount: rawRange.rowCount,
      sheetId: rawRange.sheetId,
      startColumn: rawRange.startColumn,
      startRow: rawRange.startRow,
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
      this.#recordHistoryNode(nextState);
      this.#commitState(nextState);
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

  #getEvaluationSnapshot(sheetId?: string): SheetEvaluationSnapshot {
    const sheet = getWorkbookSheet(this.#state, sheetId);
    const cachedSnapshot = this.#sheetEvaluationSnapshots.get(sheet.id);

    if (
      cachedSnapshot &&
      cachedSnapshot.workbookVersion === this.#state.version &&
      !cachedSnapshot.hasVolatileFunctions
    ) {
      return cachedSnapshot;
    }

    const nextSnapshots = evaluateWorkbook(this.#state, this.#state.version, {
      now: this.#evaluationClock(),
    });

    this.#sheetEvaluationSnapshots = nextSnapshots;
    const nextSnapshot = this.#sheetEvaluationSnapshots.get(sheet.id);

    if (!nextSnapshot) {
      throw new Error(`Sheet "${sheet.id}" was not found in the evaluation snapshot.`);
    }

    return nextSnapshot;
  }

  #getChartSheetReferences(): WorkbookChartSheetReference[] {
    return this.#state.sheets.map((sheet) => ({
      columnCount: getSheetColumnCount(sheet),
      id: sheet.id,
      rowCount: getSheetRowCount(sheet),
    }));
  }

  #commitState(nextState: WorkbookState) {
    this.#state = nextState;
    this.#sheetEvaluationSnapshots.clear();
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

function cloneRangeValues(values: string[][]): string[][] {
  return values.map((row) => [...row]);
}
