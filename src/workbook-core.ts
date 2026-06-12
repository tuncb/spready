export const DEFAULT_INITIAL_ROWS = 200;
export const DEFAULT_INITIAL_COLUMNS = 50;
export const DEFAULT_SHEET_NAME = "Sheet 1";
export const DEFAULT_COLUMN_WIDTH = 140;
export const MIN_COLUMN_WIDTH = 48;
export const MAX_COLUMN_WIDTH = 640;
export const DEFAULT_CHART_LAYOUT_WIDTH = 420;
export const DEFAULT_CHART_LAYOUT_HEIGHT = 260;
export const MIN_CHART_LAYOUT_WIDTH = 180;
export const MIN_CHART_LAYOUT_HEIGHT = 140;

export interface ControlServerInfo {
  host: string;
  port: number;
  protocol: "jsonl";
}

export interface ControlAppStatus {
  focusedWindowCount: number;
  frontendVisible: boolean;
  visibleWindowCount: number;
  windowCount: number;
}

export interface InstallerOptions {
  autoStart: boolean;
  startMenuShortcut: boolean;
}

export interface InstallerStatus {
  canManageInstalledInstance: boolean;
  currentVersion: string;
  installDirectory: string;
  installed: boolean;
  installedExecutablePath: string;
  isPackaged: boolean;
  options: InstallerOptions;
  platform: NodeJS.Platform;
  updateSupported: boolean;
}

export interface InstallerOperationResult {
  message: string;
  status: InstallerStatus;
}

export interface InstallerCheckUpdatesRequest {
  restart?: boolean;
  startUpdate?: boolean;
}

export interface InstallerCheckUpdatesResult {
  assetName?: string;
  currentVersion: string;
  latestVersion?: string;
  message: string;
  releaseUrl?: string;
  status: InstallerStatus;
  updateAvailable: boolean;
  updateStarted: boolean;
}

export interface WorkbookSheet {
  id: string;
  name: string;
  cells: string[][];
  cellStyles: Record<string, WorkbookCellStyle>;
  columnWidths: Record<string, number>;
  sourceFilePath?: string;
}

export interface WorkbookState {
  charts: WorkbookChart[];
  documentFilePath?: string;
  hasUnsavedChanges: boolean;
  version: number;
  activeSheetId: string;
  nextChartNumber: number;
  nextTableNumber: number;
  nextSheetNumber: number;
  sheets: WorkbookSheet[];
  tables: WorkbookTable[];
}

export interface SheetSummary {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  columnWidths?: Record<string, number>;
  sourceFilePath?: string;
}

export interface WorkbookSummary {
  documentFilePath?: string;
  charts: WorkbookChartSummary[];
  hasUnsavedChanges: boolean;
  version: number;
  activeSheetId: string;
  activeSheetName: string;
  sheets: SheetSummary[];
  tables: WorkbookTableSummary[];
}

export type WorkbookTableSortDirection = "ascending" | "descending";

export type WorkbookTableSortValueMode = "display" | "raw";

export interface WorkbookTableRange {
  sheetId: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

export interface WorkbookTableSortKey {
  columnIndex: number;
  direction: WorkbookTableSortDirection;
}

export interface WorkbookTableSortState {
  keys: WorkbookTableSortKey[];
  valueMode: WorkbookTableSortValueMode;
}

export interface WorkbookTable {
  hasHeaderRow: boolean;
  id: string;
  name: string;
  range: WorkbookTableRange;
  sortState?: WorkbookTableSortState;
}

export interface WorkbookTableSummary {
  hasHeaderRow: boolean;
  id: string;
  name: string;
  range: WorkbookTableRange;
  sheetId: string;
  sortState?: WorkbookTableSortState;
}

export interface WorkbookSheetTablesResult {
  sheetId: string;
  sheetName: string;
  tables: WorkbookTable[];
}

export const WORKBOOK_CHART_TYPES = ["bar", "line", "area", "pie", "scatter"] as const;

export type WorkbookChartType = (typeof WORKBOOK_CHART_TYPES)[number];

export type WorkbookChartSeriesLayout = "column" | "row";

export type WorkbookChartStatus = "ok" | "invalid";

export type WorkbookChartDimensionType = "number" | "ordinal" | "time";

export interface WorkbookChartRange {
  sheetId: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

export interface WorkbookChartSource {
  range: WorkbookChartRange;
  seriesLayoutBy: WorkbookChartSeriesLayout;
  sourceHeader: boolean;
}

export interface WorkbookChartLayout {
  height: number;
  offsetX: number;
  offsetY: number;
  startColumn: number;
  startRow: number;
  width: number;
  zIndex: number;
}

export interface WorkbookChartCartesianSpec {
  family: "cartesian";
  chartType: Exclude<WorkbookChartType, "pie">;
  source: WorkbookChartSource;
  categoryDimension: number;
  valueDimensions: number[];
  smooth?: boolean;
  stacked?: boolean;
}

export interface WorkbookChartPieSpec {
  family: "pie";
  chartType: "pie";
  source: WorkbookChartSource;
  nameDimension: number;
  valueDimension: number;
}

export type WorkbookChartSpec = WorkbookChartCartesianSpec | WorkbookChartPieSpec;

export interface WorkbookChart {
  id: string;
  layout: WorkbookChartLayout;
  name: string;
  sheetId: string;
  spec: WorkbookChartSpec;
}

export interface WorkbookChartSheetReference {
  id: string;
  rowCount: number;
  columnCount: number;
}

export type WorkbookChartValidationIssueCode =
  | "CROSS_SHEET_SOURCE"
  | "EMPTY_RANGE"
  | "EMPTY_VALUE_DIMENSIONS"
  | "INVALID_DIMENSION"
  | "INVALID_RANGE_COORDINATE"
  | "MISSING_SHEET"
  | "OUT_OF_BOUNDS"
  | "REPEATED_VALUE_DIMENSION";

export interface WorkbookChartValidationIssue {
  code: WorkbookChartValidationIssueCode;
  message: string;
}

export interface WorkbookChartSummary {
  id: string;
  layout: WorkbookChartLayout;
  name: string;
  sheetId: string;
  chartType: WorkbookChartType;
  status: WorkbookChartStatus;
}

export interface WorkbookChartPreviewDimension {
  name: string;
  type: WorkbookChartDimensionType;
}

export interface WorkbookChartPreviewDataset {
  dimensions: WorkbookChartPreviewDimension[];
  source: Array<Array<string | number | null>>;
  sourceHeader: boolean;
  seriesLayoutBy: WorkbookChartSeriesLayout;
}

export interface WorkbookChartResult {
  chart: WorkbookChart;
  status: WorkbookChartStatus;
  validationIssues: WorkbookChartValidationIssue[];
}

export interface WorkbookSheetChartsResult {
  sheetId: string;
  sheetName: string;
  charts: WorkbookChart[];
}

export interface WorkbookChartPreview extends WorkbookChartResult {
  dataset: WorkbookChartPreviewDataset;
  option: Record<string, unknown>;
  warnings: string[];
}

export interface WorkbookSheetChartPreviewsResult {
  sheetId: string;
  sheetName: string;
  previews: WorkbookChartPreview[];
}

export interface CreateChartSourceRange {
  sheetId?: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

export interface CreateChartRequest {
  categoryDimension?: number;
  chartId?: string;
  chartType: WorkbookChartType;
  dryRun?: boolean;
  expectedVersion?: number;
  layout?: WorkbookChartLayout;
  name?: string;
  nameDimension?: number;
  seriesLayoutBy?: WorkbookChartSeriesLayout;
  sheetId?: string;
  smooth?: boolean;
  sourceHeader?: boolean;
  sourceRange?: CreateChartSourceRange;
  stacked?: boolean;
  valueDimension?: number;
  valueDimensions?: number[];
}

export interface CreateChartResult extends ApplyTransactionResult {
  chart: WorkbookChartSummary;
}

export interface SheetRangeRequest {
  sheetId?: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

export type FormulaErrorCode =
  | "PARSE"
  | "REF"
  | "DIV0"
  | "VALUE"
  | "CYCLE"
  | "NAME"
  | "NUM"
  | "NA"
  | "NULL";

export interface SheetRangeResult {
  sheetId: string;
  sheetName: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  values: string[][];
}

export interface SheetDisplayRangeResult {
  sheetId: string;
  sheetName: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  values: string[][];
}

export type WorkbookCellHorizontalAlign = "center" | "left" | "right";

export interface WorkbookCellStyle {
  backgroundColor?: string;
  bold?: boolean;
  fontFamily?: string;
  fontSize?: number;
  horizontalAlign?: WorkbookCellHorizontalAlign;
  italic?: boolean;
  textColor?: string;
  wrapText?: boolean;
}

export interface WorkbookCellStylePatch {
  backgroundColor?: string | null;
  bold?: boolean | null;
  fontFamily?: string | null;
  fontSize?: number | null;
  horizontalAlign?: WorkbookCellHorizontalAlign | null;
  italic?: boolean | null;
  textColor?: string | null;
  wrapText?: boolean | null;
}

export type FormatCellsMode = "clear" | "merge" | "replace";

export interface FormatCellsRequest {
  dryRun?: boolean;
  expectedVersion?: number;
  mode?: FormatCellsMode;
  ranges: SheetRangeRequest[];
  style?: WorkbookCellStylePatch;
}

export interface SheetStyleRangeResult {
  sheetId: string;
  sheetName: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  styles: Array<Array<WorkbookCellStyle | null>>;
}

export interface CellDataRequest {
  sheetId?: string;
  rowIndex: number;
  columnIndex: number;
}

export interface CellDataResult {
  sheetId: string;
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  input: string;
  display: string;
  isFormula: boolean;
  errorCode?: FormulaErrorCode;
  style?: WorkbookCellStyle;
}

export interface UsedRangeResult {
  sheetId: string;
  sheetName: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

export type WorkbookTransactionOperation =
  | {
      type: "addSheet";
      activate?: boolean;
      columnCount?: number;
      name?: string;
      rowCount?: number;
      sheetId?: string;
    }
  | {
      chartId?: string;
      layout?: WorkbookChartLayout;
      name?: string;
      sheetId?: string;
      spec: WorkbookChartSpec;
      type: "addChart";
    }
  | {
      hasHeaderRow?: boolean;
      name?: string;
      range: Omit<WorkbookTableRange, "sheetId"> & { sheetId?: string };
      sortState?: WorkbookTableSortState;
      tableId?: string;
      type: "addTable";
    }
  | {
      type: "clearRange";
      columnCount: number;
      rowCount: number;
      sheetId?: string;
      startColumn: number;
      startRow: number;
    }
  | {
      type: "clearRangeStyle";
      columnCount: number;
      rowCount: number;
      sheetId?: string;
      startColumn: number;
      startRow: number;
    }
  | {
      type: "deleteColumns";
      columnIndex: number;
      count: number;
      sheetId?: string;
    }
  | {
      type: "deleteRows";
      count: number;
      rowIndex: number;
      sheetId?: string;
    }
  | {
      type: "deleteSheet";
      nextActiveSheetId?: string;
      sheetId: string;
    }
  | {
      chartId: string;
      type: "deleteChart";
    }
  | {
      tableId: string;
      type: "deleteTable";
    }
  | {
      type: "insertColumns";
      columnIndex: number;
      count: number;
      sheetId?: string;
    }
  | {
      type: "insertRows";
      count: number;
      rowIndex: number;
      sheetId?: string;
    }
  | {
      type: "renameSheet";
      name: string;
      sheetId?: string;
    }
  | {
      chartId: string;
      name: string;
      type: "renameChart";
    }
  | {
      name: string;
      tableId: string;
      type: "renameTable";
    }
  | {
      type: "replaceSheet";
      name?: string;
      rows: string[][];
      sheetId?: string;
      sourceFilePath?: string;
    }
  | {
      type: "replaceSheetFromCsv";
      content: string;
      name?: string;
      sheetId?: string;
      sourceFilePath?: string;
    }
  | {
      type: "resizeSheet";
      columnCount: number;
      rowCount: number;
      sheetId?: string;
    }
  | {
      range: Omit<WorkbookTableRange, "sheetId"> & { sheetId?: string };
      tableId: string;
      type: "resizeTable";
    }
  | {
      type: "setActiveSheet";
      sheetId: string;
    }
  | {
      type: "setColumnWidth";
      columnIndex: number;
      sheetId?: string;
      width: number;
    }
  | {
      chartId: string;
      layout: WorkbookChartLayout;
      type: "setChartLayout";
    }
  | {
      chartId: string;
      spec: WorkbookChartSpec;
      type: "setChartSpec";
    }
  | {
      bodyRowOrder?: number[];
      keys: WorkbookTableSortKey[];
      tableId: string;
      type: "sortTable";
      valueMode?: WorkbookTableSortValueMode;
    }
  | {
      type: "setSheetSourceFile";
      sheetId?: string;
      sourceFilePath?: string;
    }
  | {
      type: "setCell";
      columnIndex: number;
      rowIndex: number;
      sheetId?: string;
      value: string;
    }
  | {
      type: "setCellStyle";
      columnIndex: number;
      rowIndex: number;
      sheetId?: string;
      style?: WorkbookCellStyle;
    }
  | {
      type: "setRange";
      sheetId?: string;
      startColumn: number;
      startRow: number;
      values: string[][];
    }
  | {
      type: "setRangeStyle";
      columnCount: number;
      rowCount: number;
      sheetId?: string;
      startColumn: number;
      startRow: number;
      style?: WorkbookCellStyle;
    };

export interface ApplyTransactionRequest {
  dryRun?: boolean;
  expectedVersion?: number;
  operations: WorkbookTransactionOperation[];
}

export interface ApplyTransactionResult {
  changed: boolean;
  summary: WorkbookSummary;
  version: number;
}

export interface WorkbookHistoryRequest {
  expectedVersion?: number;
}

export interface WorkbookRedoRequest extends WorkbookHistoryRequest {
  nodeId?: string;
}

export interface WorkbookHistoryCheckoutRequest extends WorkbookHistoryRequest {
  nodeId: string;
}

export interface WorkbookUndoTreeNode {
  childIds: string[];
  id: string;
  isCurrent: boolean;
  isSaved: boolean;
  parentId?: string;
  summary: WorkbookSummary;
}

export interface WorkbookUndoTree {
  canRedo: boolean;
  canUndo: boolean;
  currentNodeId: string;
  nodes: WorkbookUndoTreeNode[];
  rootNodeId: string;
  savedNodeId?: string;
}

export interface WorkbookHistoryResult extends ApplyTransactionResult {
  undoTree: WorkbookUndoTree;
}

export interface ImportCsvFileRequest {
  filePath: string;
  name?: string;
  sheetId?: string;
}

export interface ExportCsvFileRequest {
  filePath: string;
  sheetId?: string;
}

export type ClipboardRangeMode = "display" | "raw";

export interface ClipboardRangePayload {
  displayText: string;
  displayValues: string[][];
  rawText: string;
  rawValues: string[][];
  styles?: ClipboardCellStyle[];
  tables?: ClipboardTable[];
}

export interface ClipboardCellStyle {
  columnOffset: number;
  rowOffset: number;
  style: WorkbookCellStyle;
}

export interface ClipboardTableRange {
  startColumnOffset: number;
  startRowOffset: number;
  rowCount: number;
  columnCount: number;
}

export interface ClipboardTableSortKey {
  columnOffset: number;
  direction: WorkbookTableSortDirection;
}

export interface ClipboardTableSortState {
  keys: ClipboardTableSortKey[];
  valueMode: WorkbookTableSortValueMode;
}

export interface ClipboardTable {
  hasHeaderRow: boolean;
  name: string;
  range: ClipboardTableRange;
  sortState?: ClipboardTableSortState;
}

export interface CopyRangeRequest extends SheetRangeRequest {
  mode?: ClipboardRangeMode;
}

export interface CopyRangeResult {
  columnCount: number;
  mode: ClipboardRangeMode;
  rowCount: number;
  sheetId: string;
  sheetName: string;
  startColumn: number;
  startRow: number;
  text: string;
  values: string[][];
  clipboard: ClipboardRangePayload;
}

export interface CutRangeRequest extends SheetRangeRequest {
  mode?: ClipboardRangeMode;
}

export interface CutRangeResult {
  changed: boolean;
  clipboard: ClipboardRangePayload;
  columnCount: number;
  mode: ClipboardRangeMode;
  rowCount: number;
  sheetId: string;
  sheetName: string;
  startColumn: number;
  startRow: number;
  summary: WorkbookSummary;
  text: string;
  values: string[][];
  version: number;
}

export interface PasteRangeRequest {
  clipboard?: ClipboardRangePayload;
  mode?: ClipboardRangeMode;
  sheetId?: string;
  startColumn: number;
  startRow: number;
  text?: string;
  values?: string[][];
}

export interface ClearRangeRequest {
  columnCount: number;
  rowCount: number;
  sheetId?: string;
  startColumn: number;
  startRow: number;
}

export interface CreateNewWorkbookRequest {
  discardUnsavedChanges?: boolean;
}

export interface OpenWorkbookFileRequest {
  discardUnsavedChanges?: boolean;
  filePath: string;
}

export interface SaveWorkbookFileRequest {
  filePath: string;
}

export interface CsvFileOperationResult {
  changed: boolean;
  filePath: string;
  summary: WorkbookSummary;
  version: number;
}

export interface WorkbookFileOperationResult {
  changed: boolean;
  filePath: string;
  summary: WorkbookSummary;
  version: number;
}

export interface WorkbookTransactionExecutionResult {
  changed: boolean;
  state: WorkbookState;
}

let nextSheetIdSequence = 1;

export function createSheet(rowCount: number, columnCount: number): string[][] {
  const normalizedRowCount = Math.max(1, Math.floor(rowCount));
  const normalizedColumnCount = Math.max(1, Math.floor(columnCount));

  return Array.from({ length: normalizedRowCount }, () => Array(normalizedColumnCount).fill(""));
}

export function normalizeSheet(rows: string[][]): string[][] {
  const rowCount = Math.max(rows.length, 1);
  const columnCount = Math.max(1, ...rows.map((row) => row.length));

  return Array.from({ length: rowCount }, (_, rowIndex) => {
    const sourceRow = rows[rowIndex] ?? [];

    return Array.from({ length: columnCount }, (_, columnIndex) => sourceRow[columnIndex] ?? "");
  });
}

export function parseCsv(content: string): string[][] {
  return parseDelimitedText(content, ",");
}

export function parseTsv(content: string): string[][] {
  return parseDelimitedText(content, "\t");
}

function parseDelimitedText(content: string, delimiter: string): string[][] {
  if (content.length === 0) {
    return normalizeSheet([]);
  }

  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentValue = "";
  let isQuoted = false;

  for (let index = 0; index < content.length; index += 1) {
    const character = content[index];

    if (isQuoted) {
      if (character === '"') {
        if (content[index + 1] === '"') {
          currentValue += '"';
          index += 1;
        } else {
          isQuoted = false;
        }
      } else {
        currentValue += character;
      }

      continue;
    }

    if (character === '"') {
      isQuoted = true;
      continue;
    }

    if (character === delimiter) {
      currentRow.push(currentValue);
      currentValue = "";
      continue;
    }

    if (character === "\n" || character === "\r") {
      if (character === "\r" && content[index + 1] === "\n") {
        index += 1;
      }

      currentRow.push(currentValue);
      rows.push(currentRow);
      currentRow = [];
      currentValue = "";
      continue;
    }

    currentValue += character;
  }

  if (currentRow.length > 0 || currentValue.length > 0) {
    currentRow.push(currentValue);
    rows.push(currentRow);
  }

  return normalizeSheet(rows);
}

export function getUsedRange(sheet: WorkbookSheet): UsedRangeResult {
  let lastRowIndex = -1;
  let lastColumnIndex = -1;

  for (let rowIndex = 0; rowIndex < sheet.cells.length; rowIndex += 1) {
    const row = sheet.cells[rowIndex];

    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (row[columnIndex] === "") {
        continue;
      }

      lastRowIndex = rowIndex;
      lastColumnIndex = Math.max(lastColumnIndex, columnIndex);
    }
  }

  return {
    columnCount: lastColumnIndex < 0 ? 0 : lastColumnIndex + 1,
    rowCount: lastRowIndex < 0 ? 0 : lastRowIndex + 1,
    sheetId: sheet.id,
    sheetName: sheet.name,
    startColumn: 0,
    startRow: 0,
  };
}

function escapeDelimitedValue(value: string, delimiter: string): string {
  if (value.includes(delimiter) || value.includes('"') || /[\r\n]/.test(value)) {
    return `"${value.replaceAll('"', '""')}"`;
  }

  return value;
}

export function serializeCsv(sheet: WorkbookSheet): string {
  const usedRange = getUsedRange(sheet);

  if (usedRange.rowCount === 0 || usedRange.columnCount === 0) {
    return "";
  }

  return sheet.cells
    .slice(0, usedRange.rowCount)
    .map((row) =>
      row
        .slice(0, usedRange.columnCount)
        .map((value) => escapeDelimitedValue(value, ","))
        .join(","),
    )
    .join("\r\n");
}

export function serializeTsv(values: readonly (readonly string[])[]): string {
  return values
    .map((row) => row.map((value) => escapeDelimitedValue(value ?? "", "\t")).join("\t"))
    .join("\r\n");
}

export function getColumnTitle(index: number): string {
  let current = index;
  let label = "";

  do {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);

  return label;
}

export function createWorkbookState(): WorkbookState {
  const defaultSheet = createWorkbookSheet(
    DEFAULT_SHEET_NAME,
    DEFAULT_INITIAL_ROWS,
    DEFAULT_INITIAL_COLUMNS,
  );

  return {
    activeSheetId: defaultSheet.id,
    charts: [],
    hasUnsavedChanges: false,
    nextChartNumber: 1,
    nextTableNumber: 1,
    nextSheetNumber: 2,
    sheets: [defaultSheet],
    tables: [],
    version: 0,
  };
}

export function getWorkbookSummary(workbook: WorkbookState): WorkbookSummary {
  const activeSheet = getSheetById(workbook, workbook.activeSheetId);
  const chartSheets: WorkbookChartSheetReference[] = workbook.sheets.map((sheet) => ({
    columnCount: getSheetColumnCount(sheet),
    id: sheet.id,
    rowCount: getSheetRowCount(sheet),
  }));

  return {
    activeSheetId: activeSheet.id,
    activeSheetName: activeSheet.name,
    charts: workbook.charts.map((chart) => createWorkbookChartSummary(chart, chartSheets)),
    documentFilePath: workbook.documentFilePath,
    hasUnsavedChanges: workbook.hasUnsavedChanges,
    sheets: workbook.sheets.map((sheet) => ({
      columnCount: getSheetColumnCount(sheet),
      columnWidths: cloneColumnWidths(sheet.columnWidths),
      id: sheet.id,
      name: sheet.name,
      rowCount: getSheetRowCount(sheet),
      sourceFilePath: sheet.sourceFilePath,
    })),
    tables: workbook.tables.map(createWorkbookTableSummary),
    version: workbook.version,
  };
}

export function normalizeWorkbookSheetName(name: string): string {
  const normalizedName = name.trim();

  if (normalizedName.length === 0) {
    throw new Error("Sheet name is required.");
  }

  return normalizedName;
}

export function assertWorkbookSheetNamesAreUnique(workbook: {
  sheets: readonly { name: string }[];
}): void {
  const seenNames = new Map<string, string>();

  for (const sheet of workbook.sheets) {
    const normalizedName = normalizeWorkbookSheetName(sheet.name);
    const nameKey = getWorkbookSheetNameKey(normalizedName);
    const existingName = seenNames.get(nameKey);

    if (existingName) {
      throw new Error(
        `Sheet name "${normalizedName}" already exists as "${existingName}". Sheet names must be unique case-insensitively.`,
      );
    }

    seenNames.set(nameKey, normalizedName);
  }
}

export function cloneWorkbookChart(chart: WorkbookChart): WorkbookChart {
  return {
    ...chart,
    layout: cloneWorkbookChartLayout(chart.layout),
    spec: cloneWorkbookChartSpec(chart.spec),
  };
}

export function cloneWorkbookTable(table: WorkbookTable): WorkbookTable {
  return {
    hasHeaderRow: table.hasHeaderRow,
    id: table.id,
    name: table.name,
    range: { ...table.range },
    ...(table.sortState ? { sortState: cloneWorkbookTableSortState(table.sortState) } : {}),
  };
}

export function getWorkbookTableById(workbook: WorkbookState, tableId: string): WorkbookTable {
  const table = workbook.tables.find((entry) => entry.id === tableId);

  if (!table) {
    throw new Error(`Table "${tableId}" was not found.`);
  }

  return table;
}

export function getWorkbookSheetTables(
  workbook: WorkbookState,
  sheetId?: string,
): WorkbookSheetTablesResult {
  const sheet = getWorkbookSheet(workbook, sheetId ?? workbook.activeSheetId);

  return {
    sheetId: sheet.id,
    sheetName: sheet.name,
    tables: workbook.tables
      .filter((table) => table.range.sheetId === sheet.id)
      .map((table) => cloneWorkbookTable(table)),
  };
}

export function createClipboardRangePayload(
  workbook: WorkbookState,
  rawRange: SheetRangeResult,
  displayRange: SheetDisplayRangeResult,
): ClipboardRangePayload {
  const sheet = getSheetById(workbook, rawRange.sheetId);
  const selectionRange = createWorkbookRangeFromSheetRange(rawRange);
  const styles: ClipboardCellStyle[] = [];
  const tables: ClipboardTable[] = [];

  for (const [key, style] of Object.entries(sheet.cellStyles)) {
    const [rowText, columnText] = key.split(":");
    const rowIndex = Number.parseInt(rowText, 10);
    const columnIndex = Number.parseInt(columnText, 10);

    if (
      rowIndex < rawRange.startRow ||
      rowIndex >= rawRange.startRow + rawRange.rowCount ||
      columnIndex < rawRange.startColumn ||
      columnIndex >= rawRange.startColumn + rawRange.columnCount
    ) {
      continue;
    }

    styles.push({
      columnOffset: columnIndex - rawRange.startColumn,
      rowOffset: rowIndex - rawRange.startRow,
      style: cloneWorkbookCellStyle(style),
    });
  }

  for (const table of workbook.tables) {
    if (!workbookTableRangeContains(selectionRange, table.range)) {
      continue;
    }

    tables.push({
      hasHeaderRow: table.hasHeaderRow,
      name: table.name,
      range: {
        columnCount: table.range.columnCount,
        rowCount: table.range.rowCount,
        startColumnOffset: table.range.startColumn - rawRange.startColumn,
        startRowOffset: table.range.startRow - rawRange.startRow,
      },
      ...(table.sortState
        ? {
            sortState: {
              keys: table.sortState.keys.map((key) => ({
                columnOffset: key.columnIndex - table.range.startColumn,
                direction: key.direction,
              })),
              valueMode: table.sortState.valueMode,
            },
          }
        : {}),
    });
  }

  return {
    displayText: serializeTsv(displayRange.values),
    displayValues: cloneRangeValues(displayRange.values),
    rawText: serializeTsv(rawRange.values),
    rawValues: cloneRangeValues(rawRange.values),
    styles,
    tables,
  };
}

export function buildCutRangeOperations(
  workbook: WorkbookState,
  range: SheetRangeResult,
): WorkbookTransactionOperation[] {
  const targetRange = createWorkbookRangeFromSheetRange(range);
  const operations: WorkbookTransactionOperation[] = [
    {
      columnCount: range.columnCount,
      rowCount: range.rowCount,
      sheetId: range.sheetId,
      startColumn: range.startColumn,
      startRow: range.startRow,
      type: "clearRange",
    },
    {
      columnCount: range.columnCount,
      rowCount: range.rowCount,
      sheetId: range.sheetId,
      startColumn: range.startColumn,
      startRow: range.startRow,
      type: "clearRangeStyle",
    },
  ];

  for (const table of workbook.tables) {
    if (table.range.sheetId !== range.sheetId || !rangesOverlap(table.range, targetRange)) {
      continue;
    }

    if (
      workbookTableRangeContains(targetRange, table.range) ||
      (table.hasHeaderRow && workbookRangeIntersectsTableHeader(targetRange, table))
    ) {
      operations.push({
        tableId: table.id,
        type: "deleteTable",
      });
      continue;
    }

    if (table.sortState) {
      operations.push({
        range: { ...table.range },
        tableId: table.id,
        type: "resizeTable",
      });
    }
  }

  return operations;
}

export function buildPasteRangeOperations(
  workbook: WorkbookState,
  request: PasteRangeRequest,
  values: string[][],
): WorkbookTransactionOperation[] {
  if (values.length === 0) {
    return [];
  }

  const columnCount = Math.max(0, ...values.map((row) => row.length));

  if (columnCount === 0) {
    return [];
  }

  const sheet = getWorkbookSheet(workbook, request.sheetId);
  const targetRange: WorkbookTableRange = {
    columnCount,
    rowCount: values.length,
    sheetId: sheet.id,
    startColumn: request.startColumn,
    startRow: request.startRow,
  };
  const operations: WorkbookTransactionOperation[] = [];
  const clipboardTables = request.clipboard?.tables ?? [];
  const replacedTableIds = new Set<string>();

  if (clipboardTables.length > 0) {
    const targetTableRanges = clipboardTables.map((table) =>
      createWorkbookRangeFromClipboardTable(request, sheet.id, table),
    );
    const replacedTables = workbook.tables.filter(
      (table) =>
        table.range.sheetId === sheet.id &&
        targetTableRanges.some((tableRange) => rangesOverlap(table.range, tableRange)),
    );

    for (const table of replacedTables) {
      replacedTableIds.add(table.id);
      operations.push({
        tableId: table.id,
        type: "deleteTable",
      });
      operations.push({
        columnCount: table.range.columnCount,
        rowCount: table.range.rowCount,
        sheetId: sheet.id,
        startColumn: table.range.startColumn,
        startRow: table.range.startRow,
        type: "clearRange",
      });
      operations.push({
        columnCount: table.range.columnCount,
        rowCount: table.range.rowCount,
        sheetId: sheet.id,
        startColumn: table.range.startColumn,
        startRow: table.range.startRow,
        type: "clearRangeStyle",
      });
    }
  } else {
    const targetTopLeftTable = workbook.tables.find(
      (table) =>
        table.range.sheetId === sheet.id &&
        table.range.startRow === request.startRow &&
        table.range.startColumn === request.startColumn,
    );

    if (targetTopLeftTable) {
      for (const remainder of getWorkbookRangeRemainders(targetTopLeftTable.range, targetRange)) {
        operations.push({
          columnCount: remainder.columnCount,
          rowCount: remainder.rowCount,
          sheetId: sheet.id,
          startColumn: remainder.startColumn,
          startRow: remainder.startRow,
          type: "clearRange",
        });
        operations.push({
          columnCount: remainder.columnCount,
          rowCount: remainder.rowCount,
          sheetId: sheet.id,
          startColumn: remainder.startColumn,
          startRow: remainder.startRow,
          type: "clearRangeStyle",
        });
      }

      operations.push({
        range: targetRange,
        tableId: targetTopLeftTable.id,
        type: "resizeTable",
      });
    }
  }

  operations.push({
    sheetId: sheet.id,
    startColumn: request.startColumn,
    startRow: request.startRow,
    type: "setRange",
    values,
  });

  if (request.clipboard?.styles) {
    operations.push({
      columnCount: targetRange.columnCount,
      rowCount: targetRange.rowCount,
      sheetId: sheet.id,
      startColumn: targetRange.startColumn,
      startRow: targetRange.startRow,
      type: "clearRangeStyle",
    });

    for (const cellStyle of request.clipboard.styles) {
      if (
        cellStyle.rowOffset < 0 ||
        cellStyle.columnOffset < 0 ||
        cellStyle.rowOffset >= targetRange.rowCount ||
        cellStyle.columnOffset >= targetRange.columnCount
      ) {
        continue;
      }

      operations.push({
        columnIndex: targetRange.startColumn + cellStyle.columnOffset,
        rowIndex: targetRange.startRow + cellStyle.rowOffset,
        sheetId: sheet.id,
        style: cloneWorkbookCellStyle(cellStyle.style),
        type: "setCellStyle",
      });
    }
  }

  if (clipboardTables.length > 0) {
    const usedNames = new Set(
      workbook.tables
        .filter((table) => !replacedTableIds.has(table.id))
        .map((table) => getWorkbookTableNameKey(table.name)),
    );

    for (const table of clipboardTables) {
      const tableRange = createWorkbookRangeFromClipboardTable(request, sheet.id, table);
      const tableName = getAvailablePastedWorkbookTableName(usedNames, table.name);

      operations.push({
        hasHeaderRow: table.hasHeaderRow,
        name: tableName,
        range: tableRange,
        ...(table.sortState
          ? {
              sortState: {
                keys: table.sortState.keys.map((key) => ({
                  columnIndex: tableRange.startColumn + key.columnOffset,
                  direction: key.direction,
                })),
                valueMode: table.sortState.valueMode,
              },
            }
          : {}),
        type: "addTable",
      });
    }
  }

  return operations;
}

export function getWorkbookChartById(workbook: WorkbookState, chartId: string): WorkbookChart {
  const chart = workbook.charts.find((entry) => entry.id === chartId);

  if (!chart) {
    throw new Error(`Chart "${chartId}" was not found.`);
  }

  return chart;
}

export function getWorkbookSheetCharts(
  workbook: WorkbookState,
  sheetId?: string,
): WorkbookSheetChartsResult {
  const sheet = getSheetById(workbook, sheetId ?? workbook.activeSheetId);

  return {
    charts: workbook.charts
      .filter((chart) => chart.sheetId === sheet.id)
      .map((chart) => cloneWorkbookChart(chart)),
    sheetId: sheet.id,
    sheetName: sheet.name,
  };
}

export function buildCreateChartOperation(
  workbook: WorkbookState,
  request: CreateChartRequest,
): {
  chartId: string;
  operation: Extract<WorkbookTransactionOperation, { type: "addChart" }>;
} {
  const sheet = getWorkbookSheet(workbook, request.sheetId);
  const sourceSheet = getWorkbookSheet(workbook, request.sourceRange?.sheetId ?? request.sheetId);
  const chartId =
    normalizeOptionalChartId(request.chartId) ?? createChartId(workbook.nextChartNumber);
  const range = createChartRangeFromRequest(sourceSheet, request.sourceRange);
  const seriesLayoutBy = request.seriesLayoutBy ?? "column";
  const dimensionCount = seriesLayoutBy === "row" ? range.rowCount : range.columnCount;
  const source: WorkbookChartSource = {
    range,
    seriesLayoutBy,
    sourceHeader: request.sourceHeader ?? true,
  };
  const spec: WorkbookChartSpec =
    request.chartType === "pie"
      ? {
          chartType: "pie",
          family: "pie",
          nameDimension: request.nameDimension ?? 0,
          source,
          valueDimension:
            request.valueDimension ?? getDefaultSecondaryChartDimension(0, dimensionCount),
        }
      : {
          categoryDimension: request.categoryDimension ?? 0,
          chartType: request.chartType,
          family: "cartesian",
          ...(request.chartType === "line" || request.chartType === "area"
            ? { smooth: request.smooth ?? false }
            : {}),
          source,
          ...(request.chartType === "bar" || request.chartType === "area"
            ? { stacked: request.stacked ?? false }
            : {}),
          valueDimensions:
            request.valueDimensions && request.valueDimensions.length > 0
              ? [...request.valueDimensions]
              : getDefaultValueChartDimensions(request.categoryDimension ?? 0, dimensionCount),
        };

  return {
    chartId,
    operation: {
      chartId,
      layout: request.layout,
      name: request.name,
      sheetId: sheet.id,
      spec,
      type: "addChart",
    },
  };
}

export function buildFormatCellsOperations(
  workbook: WorkbookState,
  request: FormatCellsRequest,
): WorkbookTransactionOperation[] {
  const mode = request.mode ?? "merge";
  const ranges = request.ranges.map((range) => normalizeFormatCellsRange(workbook, range));

  if (mode === "clear") {
    return ranges.map((range) => ({
      ...range,
      type: "clearRangeStyle",
    }));
  }

  if (mode === "replace") {
    const style = patchWorkbookCellStyle(undefined, request.style);

    return ranges.map((range) => ({
      ...range,
      style,
      type: "setRangeStyle",
    }));
  }

  if (!request.style) {
    return [];
  }

  const operations: WorkbookTransactionOperation[] = [];

  for (const range of ranges) {
    const sheet = getWorkbookSheet(workbook, range.sheetId);

    for (let rowIndex = range.startRow; rowIndex < range.startRow + range.rowCount; rowIndex += 1) {
      for (
        let columnIndex = range.startColumn;
        columnIndex < range.startColumn + range.columnCount;
        columnIndex += 1
      ) {
        const currentStyle = sheet.cellStyles[getCellKey(rowIndex, columnIndex)];
        const style = patchWorkbookCellStyle(currentStyle, request.style);

        if (workbookCellStylesEqual(currentStyle, style)) {
          continue;
        }

        operations.push({
          columnIndex,
          rowIndex,
          sheetId: sheet.id,
          style,
          type: "setCellStyle",
        });
      }
    }
  }

  return operations;
}

export function getSheetRange(
  workbook: WorkbookState,
  request: SheetRangeRequest,
): SheetRangeResult {
  const sheet = getSheetById(workbook, request.sheetId ?? workbook.activeSheetId);
  const rowCount = getSheetRowCount(sheet);
  const columnCount = getSheetColumnCount(sheet);
  const startRow = clampToRange(request.startRow, 0, rowCount);
  const startColumn = clampToRange(request.startColumn, 0, columnCount);
  const requestedRowCount = Math.max(0, Math.floor(request.rowCount));
  const requestedColumnCount = Math.max(0, Math.floor(request.columnCount));
  const boundedRowCount = Math.max(0, Math.min(requestedRowCount, rowCount - startRow));
  const boundedColumnCount = Math.max(0, Math.min(requestedColumnCount, columnCount - startColumn));

  return {
    columnCount: boundedColumnCount,
    rowCount: boundedRowCount,
    sheetId: sheet.id,
    sheetName: sheet.name,
    startColumn,
    startRow,
    values: Array.from({ length: boundedRowCount }, (_, rowOffset) => {
      const row = sheet.cells[startRow + rowOffset] ?? [];

      return Array.from(
        { length: boundedColumnCount },
        (_, columnOffset) => row[startColumn + columnOffset] ?? "",
      );
    }),
  };
}

export function getSheetStyleRange(
  workbook: WorkbookState,
  request: SheetRangeRequest,
): SheetStyleRangeResult {
  const sheet = getSheetById(workbook, request.sheetId ?? workbook.activeSheetId);
  const rowCount = getSheetRowCount(sheet);
  const columnCount = getSheetColumnCount(sheet);
  const startRow = clampToRange(request.startRow, 0, rowCount);
  const startColumn = clampToRange(request.startColumn, 0, columnCount);
  const requestedRowCount = Math.max(0, Math.floor(request.rowCount));
  const requestedColumnCount = Math.max(0, Math.floor(request.columnCount));
  const boundedRowCount = Math.max(0, Math.min(requestedRowCount, rowCount - startRow));
  const boundedColumnCount = Math.max(0, Math.min(requestedColumnCount, columnCount - startColumn));

  return {
    columnCount: boundedColumnCount,
    rowCount: boundedRowCount,
    sheetId: sheet.id,
    sheetName: sheet.name,
    startColumn,
    startRow,
    styles: Array.from({ length: boundedRowCount }, (_, rowOffset) =>
      Array.from({ length: boundedColumnCount }, (_, columnOffset) => {
        const style =
          sheet.cellStyles[getCellKey(startRow + rowOffset, startColumn + columnOffset)];

        return style ? cloneWorkbookCellStyle(style) : null;
      }),
    ),
  };
}

export function getSheetCsv(workbook: WorkbookState, sheetId?: string): string {
  return serializeCsv(getSheetById(workbook, sheetId ?? workbook.activeSheetId));
}

export function getSheetUsedRange(workbook: WorkbookState, sheetId?: string): UsedRangeResult {
  return getUsedRange(getSheetById(workbook, sheetId ?? workbook.activeSheetId));
}

export function getWorkbookSheet(workbook: WorkbookState, sheetId?: string): WorkbookSheet {
  return getSheetById(workbook, sheetId ?? workbook.activeSheetId);
}

export function isFormulaInput(value: string): boolean {
  return value.startsWith("=");
}

export function parseCellReference(reference: string): {
  rowIndex: number;
  columnIndex: number;
} {
  const match = /^([A-Za-z]+)([1-9][0-9]*)$/.exec(reference);

  if (!match) {
    throw new Error(`Invalid cell reference "${reference}".`);
  }

  const [, columnLabel, rowLabel] = match;
  let columnIndex = 0;

  for (const character of columnLabel.toUpperCase()) {
    columnIndex = columnIndex * 26 + (character.charCodeAt(0) - 64);
  }

  return {
    columnIndex: columnIndex - 1,
    rowIndex: Number.parseInt(rowLabel, 10) - 1,
  };
}

export function getWorkbookChartDimensionCount(chart: Pick<WorkbookChart, "spec">): number {
  const { range, seriesLayoutBy } = chart.spec.source;

  return seriesLayoutBy === "row" ? range.rowCount : range.columnCount;
}

export function getWorkbookChartValidationIssues(
  chart: WorkbookChart,
  sheets: readonly WorkbookChartSheetReference[],
): WorkbookChartValidationIssue[] {
  const issues: WorkbookChartValidationIssue[] = [];
  const { range } = chart.spec.source;
  const sourceSheet = sheets.find((sheet) => sheet.id === range.sheetId);
  const chartSheet = sheets.find((sheet) => sheet.id === chart.sheetId);

  if (
    !Number.isInteger(range.startRow) ||
    !Number.isInteger(range.startColumn) ||
    !Number.isInteger(range.rowCount) ||
    !Number.isInteger(range.columnCount) ||
    range.startRow < 0 ||
    range.startColumn < 0 ||
    range.rowCount < 0 ||
    range.columnCount < 0
  ) {
    issues.push({
      code: "INVALID_RANGE_COORDINATE",
      message: "Chart source ranges require non-negative integer coordinates and sizes.",
    });
  }

  if (range.rowCount < 1 || range.columnCount < 1) {
    issues.push({
      code: "EMPTY_RANGE",
      message: "Chart source ranges must contain at least one row and one column.",
    });
  }

  if (!chartSheet || !sourceSheet) {
    issues.push({
      code: "MISSING_SHEET",
      message: "Chart references a sheet that is missing from the workbook.",
    });
  } else if (
    range.startRow + range.rowCount > sourceSheet.rowCount ||
    range.startColumn + range.columnCount > sourceSheet.columnCount
  ) {
    issues.push({
      code: "OUT_OF_BOUNDS",
      message: "Chart source range extends beyond the bounds of its sheet.",
    });
  }

  const dimensionCount = getWorkbookChartDimensionCount(chart);

  if (chart.spec.family === "cartesian") {
    if (chart.spec.valueDimensions.length === 0) {
      issues.push({
        code: "EMPTY_VALUE_DIMENSIONS",
        message: "Cartesian charts require at least one value dimension in the shared contract.",
      });
    }

    if (new Set(chart.spec.valueDimensions).size !== chart.spec.valueDimensions.length) {
      issues.push({
        code: "REPEATED_VALUE_DIMENSION",
        message: "Cartesian chart value dimensions must not repeat.",
      });
    }

    const dimensionsToCheck = [chart.spec.categoryDimension, ...chart.spec.valueDimensions];

    if (
      dimensionsToCheck.some(
        (dimension) => !Number.isInteger(dimension) || dimension < 0 || dimension >= dimensionCount,
      )
    ) {
      issues.push({
        code: "INVALID_DIMENSION",
        message: "Cartesian chart dimensions must resolve inside the source table.",
      });
    }
  } else if (
    !Number.isInteger(chart.spec.nameDimension) ||
    chart.spec.nameDimension < 0 ||
    chart.spec.nameDimension >= dimensionCount ||
    !Number.isInteger(chart.spec.valueDimension) ||
    chart.spec.valueDimension < 0 ||
    chart.spec.valueDimension >= dimensionCount
  ) {
    issues.push({
      code: "INVALID_DIMENSION",
      message: "Pie chart dimensions must resolve inside the source table.",
    });
  }

  return issues;
}

export function getWorkbookChartStatus(
  chart: WorkbookChart,
  sheets: readonly WorkbookChartSheetReference[],
): WorkbookChartStatus {
  return getWorkbookChartValidationIssues(chart, sheets).length === 0 ? "ok" : "invalid";
}

export function createWorkbookChartSummary(
  chart: WorkbookChart,
  sheets: readonly WorkbookChartSheetReference[],
): WorkbookChartSummary {
  return {
    chartType: chart.spec.chartType,
    id: chart.id,
    layout: cloneWorkbookChartLayout(chart.layout),
    name: chart.name,
    sheetId: chart.sheetId,
    status: getWorkbookChartStatus(chart, sheets),
  };
}

export function createWorkbookTableSummary(table: WorkbookTable): WorkbookTableSummary {
  return {
    hasHeaderRow: table.hasHeaderRow,
    id: table.id,
    name: table.name,
    range: { ...table.range },
    sheetId: table.range.sheetId,
    ...(table.sortState ? { sortState: cloneWorkbookTableSortState(table.sortState) } : {}),
  };
}

export function adjustWorkbookTableForInsertedRows(
  table: WorkbookTable,
  sheetId: string,
  rowIndex: number,
  count: number,
): WorkbookTable {
  assertNonNegativeIndex(rowIndex, "Inserted row index");
  assertPositiveCount(count, "Inserted row count");

  if (table.range.sheetId !== sheetId) {
    return table;
  }

  const rangeEnd = table.range.startRow + table.range.rowCount;

  if (rowIndex <= table.range.startRow) {
    return updateWorkbookTableRange(table, {
      ...table.range,
      startRow: table.range.startRow + count,
    });
  }

  if (rowIndex < rangeEnd) {
    return updateWorkbookTableRange(table, {
      ...table.range,
      rowCount: table.range.rowCount + count,
    });
  }

  return table;
}

export function adjustWorkbookTableForDeletedRows(
  table: WorkbookTable,
  sheetId: string,
  rowIndex: number,
  count: number,
): WorkbookTable[] {
  assertNonNegativeIndex(rowIndex, "Deleted row index");
  assertPositiveCount(count, "Deleted row count");

  if (table.range.sheetId !== sheetId) {
    return [table];
  }

  const nextRangeAxis = adjustRangeAxisForDeletion(
    table.range.startRow,
    table.range.rowCount,
    rowIndex,
    count,
  );

  if (nextRangeAxis.count === 0) {
    return [];
  }

  return [
    updateWorkbookTableRange(table, {
      ...table.range,
      rowCount: nextRangeAxis.count,
      startRow: nextRangeAxis.start,
    }),
  ];
}

export function adjustWorkbookTableForInsertedColumns(
  table: WorkbookTable,
  sheetId: string,
  columnIndex: number,
  count: number,
): WorkbookTable {
  assertNonNegativeIndex(columnIndex, "Inserted column index");
  assertPositiveCount(count, "Inserted column count");

  if (table.range.sheetId !== sheetId) {
    return table;
  }

  const rangeEnd = table.range.startColumn + table.range.columnCount;

  if (columnIndex <= table.range.startColumn) {
    return updateWorkbookTableRange(table, {
      ...table.range,
      startColumn: table.range.startColumn + count,
    });
  }

  if (columnIndex < rangeEnd) {
    return updateWorkbookTableRange(table, {
      ...table.range,
      columnCount: table.range.columnCount + count,
    });
  }

  return table;
}

export function adjustWorkbookTableForDeletedColumns(
  table: WorkbookTable,
  sheetId: string,
  columnIndex: number,
  count: number,
): WorkbookTable[] {
  assertNonNegativeIndex(columnIndex, "Deleted column index");
  assertPositiveCount(count, "Deleted column count");

  if (table.range.sheetId !== sheetId) {
    return [table];
  }

  const nextRangeAxis = adjustRangeAxisForDeletion(
    table.range.startColumn,
    table.range.columnCount,
    columnIndex,
    count,
  );

  if (nextRangeAxis.count === 0) {
    return [];
  }

  return [
    updateWorkbookTableRange(table, {
      ...table.range,
      columnCount: nextRangeAxis.count,
      startColumn: nextRangeAxis.start,
    }),
  ];
}

export function adjustWorkbookChartForInsertedRows(
  chart: WorkbookChart,
  sheetId: string,
  rowIndex: number,
  count: number,
): WorkbookChart {
  assertNonNegativeIndex(rowIndex, "Inserted row index");
  assertPositiveCount(count, "Inserted row count");

  if (chart.spec.source.range.sheetId !== sheetId) {
    return chart;
  }

  const range = chart.spec.source.range;
  const rangeEnd = range.startRow + range.rowCount;

  if (rowIndex <= range.startRow) {
    return updateWorkbookChartRange(chart, {
      ...range,
      startRow: range.startRow + count,
    });
  }

  if (rowIndex < rangeEnd) {
    return updateWorkbookChartRange(chart, {
      ...range,
      rowCount: range.rowCount + count,
    });
  }

  return chart;
}

export function adjustWorkbookChartForDeletedRows(
  chart: WorkbookChart,
  sheetId: string,
  rowIndex: number,
  count: number,
): WorkbookChart {
  assertNonNegativeIndex(rowIndex, "Deleted row index");
  assertPositiveCount(count, "Deleted row count");

  if (chart.spec.source.range.sheetId !== sheetId) {
    return chart;
  }

  const range = chart.spec.source.range;
  const nextRangeAxis = adjustRangeAxisForDeletion(range.startRow, range.rowCount, rowIndex, count);

  return updateWorkbookChartRange(chart, {
    ...range,
    rowCount: nextRangeAxis.count,
    startRow: nextRangeAxis.start,
  });
}

export function adjustWorkbookChartForInsertedColumns(
  chart: WorkbookChart,
  sheetId: string,
  columnIndex: number,
  count: number,
): WorkbookChart {
  assertNonNegativeIndex(columnIndex, "Inserted column index");
  assertPositiveCount(count, "Inserted column count");

  if (chart.spec.source.range.sheetId !== sheetId) {
    return chart;
  }

  const range = chart.spec.source.range;
  const rangeEnd = range.startColumn + range.columnCount;

  if (columnIndex <= range.startColumn) {
    return updateWorkbookChartRange(chart, {
      ...range,
      startColumn: range.startColumn + count,
    });
  }

  if (columnIndex < rangeEnd) {
    return updateWorkbookChartRange(chart, {
      ...range,
      columnCount: range.columnCount + count,
    });
  }

  return chart;
}

export function adjustWorkbookChartForDeletedColumns(
  chart: WorkbookChart,
  sheetId: string,
  columnIndex: number,
  count: number,
): WorkbookChart {
  assertNonNegativeIndex(columnIndex, "Deleted column index");
  assertPositiveCount(count, "Deleted column count");

  if (chart.spec.source.range.sheetId !== sheetId) {
    return chart;
  }

  const range = chart.spec.source.range;
  const nextRangeAxis = adjustRangeAxisForDeletion(
    range.startColumn,
    range.columnCount,
    columnIndex,
    count,
  );

  return updateWorkbookChartRange(chart, {
    ...range,
    columnCount: nextRangeAxis.count,
    startColumn: nextRangeAxis.start,
  });
}

export function applyWorkbookTransaction(
  previousState: WorkbookState,
  request: ApplyTransactionRequest,
): WorkbookTransactionExecutionResult {
  if (request.operations.length === 0) {
    return {
      changed: false,
      state: previousState,
    };
  }

  const nextState: WorkbookState = {
    ...previousState,
    sheets: [...previousState.sheets],
    tables: [...previousState.tables],
  };
  const clonedSheetIds = new Set<string>();
  let changed = false;

  for (const operation of request.operations) {
    switch (operation.type) {
      case "addSheet": {
        const nextSheetName =
          operation.name === undefined
            ? getNextAvailableWorkbookSheetName(nextState)
            : {
                name: assertWorkbookSheetNameAvailable(nextState, operation.name),
                nextSheetNumber: nextState.nextSheetNumber + 1,
              };
        const sheetName = nextSheetName.name;
        const sheetId = operation.sheetId?.trim() || createSheetId();

        if (findSheetIndex(nextState, sheetId) >= 0) {
          throw new Error(`Sheet "${sheetId}" already exists.`);
        }

        nextState.sheets.push(
          createWorkbookSheet(
            sheetName,
            operation.rowCount ?? DEFAULT_INITIAL_ROWS,
            operation.columnCount ?? DEFAULT_INITIAL_COLUMNS,
            sheetId,
          ),
        );
        nextState.nextSheetNumber = nextSheetName.nextSheetNumber;

        if (operation.activate ?? true) {
          nextState.activeSheetId = sheetId;
        }

        changed = true;
        break;
      }

      case "addChart": {
        const chartId =
          normalizeOptionalChartId(operation.chartId) ?? createChartId(nextState.nextChartNumber);

        if (findChartIndex(nextState, chartId) >= 0) {
          throw new Error(`Chart "${chartId}" already exists.`);
        }

        const chart = createWorkbookChart(
          chartId,
          operation.name?.trim() || `Chart ${nextState.nextChartNumber}`,
          operation.spec,
          operation.layout,
          operation.sheetId ?? operation.spec.source.range.sheetId,
          nextState,
        );

        assertCreatableWorkbookChart(chart, nextState, "added");
        assertWorkbookChartLayoutInBounds(chart, nextState, "added");

        nextState.charts = [...nextState.charts, chart];
        nextState.nextChartNumber += 1;
        changed = true;
        break;
      }

      case "addTable": {
        const range = normalizeWorkbookTableRange(nextState, operation.range);
        const tableId =
          normalizeOptionalTableId(operation.tableId) ?? createTableId(nextState.nextTableNumber);
        const nextTableName =
          operation.name === undefined
            ? getNextAvailableWorkbookTableName(nextState)
            : {
                name: assertWorkbookTableNameAvailable(nextState, operation.name),
                nextTableNumber: nextState.nextTableNumber + 1,
              };

        if (findTableIndex(nextState, tableId) >= 0) {
          throw new Error(`Table "${tableId}" already exists.`);
        }

        const table: WorkbookTable = {
          hasHeaderRow: operation.hasHeaderRow ?? true,
          id: tableId,
          name: nextTableName.name,
          range,
        };
        const sortState =
          operation.sortState === undefined
            ? undefined
            : normalizeWorkbookTableSortState(
                table,
                operation.sortState.keys,
                operation.sortState.valueMode,
              );
        const nextTable: WorkbookTable = {
          ...table,
          ...(sortState ? { sortState } : {}),
        };

        assertWorkbookTableRangeAvailable(nextState, nextTable);
        nextState.tables = [...nextState.tables, nextTable];
        nextState.nextTableNumber = Math.max(
          nextTableName.nextTableNumber,
          getNextWorkbookTableNumberForId(tableId),
        );
        changed = true;
        break;
      }

      case "clearRange": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const maxRow = getSheetRowCount(sheet);
        const maxColumn = getSheetColumnCount(sheet);
        const startRow = clampToRange(operation.startRow, 0, maxRow);
        const startColumn = clampToRange(operation.startColumn, 0, maxColumn);
        const endRow = Math.min(maxRow, startRow + Math.max(0, Math.floor(operation.rowCount)));
        const endColumn = Math.min(
          maxColumn,
          startColumn + Math.max(0, Math.floor(operation.columnCount)),
        );
        let rangeChanged = false;

        for (let rowIndex = startRow; rowIndex < endRow; rowIndex += 1) {
          const row = sheet.cells[rowIndex];

          for (let columnIndex = startColumn; columnIndex < endColumn; columnIndex += 1) {
            if (row[columnIndex] === "") {
              continue;
            }

            row[columnIndex] = "";
            rangeChanged = true;
            changed = true;
          }
        }

        if (rangeChanged) {
          clearWorkbookTableSortStatesInRange(nextState, {
            columnCount: endColumn - startColumn,
            rowCount: endRow - startRow,
            sheetId: sheet.id,
            startColumn,
            startRow,
          });
        }

        break;
      }

      case "clearRangeStyle": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);

        if (
          clearCellStylesInRange(
            sheet,
            operation.startRow,
            operation.startColumn,
            operation.rowCount,
            operation.columnCount,
          )
        ) {
          changed = true;
        }

        break;
      }

      case "deleteColumns": {
        assertPositiveCount(operation.count, "Column delete count");

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const currentColumnCount = getSheetColumnCount(sheet);
        const deleteStart = clampToRange(operation.columnIndex, 0, currentColumnCount);
        const requestedDeleteCount = Math.min(operation.count, currentColumnCount - deleteStart);

        if (requestedDeleteCount === 0) {
          break;
        }

        if (requestedDeleteCount >= currentColumnCount) {
          for (const row of sheet.cells) {
            row.splice(0, row.length, "");
          }
        } else {
          for (const row of sheet.cells) {
            row.splice(deleteStart, requestedDeleteCount);
          }
        }

        sheet.cellStyles = deleteColumnStyles(sheet.cellStyles, deleteStart, requestedDeleteCount);
        sheet.columnWidths = deleteColumnWidths(
          sheet.columnWidths,
          deleteStart,
          requestedDeleteCount,
        );
        nextState.tables = nextState.tables.flatMap((table) =>
          adjustWorkbookTableForDeletedColumns(table, sheet.id, deleteStart, requestedDeleteCount),
        );

        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartLayoutForDeletedColumns(
            adjustWorkbookChartForDeletedColumns(
              chart,
              sheet.id,
              deleteStart,
              requestedDeleteCount,
            ),
            sheet.id,
            deleteStart,
            requestedDeleteCount,
            getSheetColumnCount(sheet),
          ),
        );
        changed = true;
        break;
      }

      case "deleteRows": {
        assertPositiveCount(operation.count, "Row delete count");

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const currentRowCount = getSheetRowCount(sheet);
        const deleteStart = clampToRange(operation.rowIndex, 0, currentRowCount);
        const requestedDeleteCount = Math.min(operation.count, currentRowCount - deleteStart);

        if (requestedDeleteCount === 0) {
          break;
        }

        if (requestedDeleteCount >= currentRowCount) {
          sheet.cells.splice(0, sheet.cells.length, Array(getSheetColumnCount(sheet)).fill(""));
        } else {
          sheet.cells.splice(deleteStart, requestedDeleteCount);
        }

        sheet.cellStyles = deleteRowStyles(sheet.cellStyles, deleteStart, requestedDeleteCount);
        nextState.tables = nextState.tables.flatMap((table) =>
          adjustWorkbookTableForDeletedRows(table, sheet.id, deleteStart, requestedDeleteCount),
        );

        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartLayoutForDeletedRows(
            adjustWorkbookChartForDeletedRows(chart, sheet.id, deleteStart, requestedDeleteCount),
            sheet.id,
            deleteStart,
            requestedDeleteCount,
            getSheetRowCount(sheet),
          ),
        );
        changed = true;
        break;
      }

      case "deleteSheet": {
        if (nextState.sheets.length === 1) {
          throw new Error("The last sheet cannot be deleted.");
        }

        const deleteIndex = findSheetIndex(nextState, operation.sheetId);

        if (deleteIndex < 0) {
          throw new Error(`Sheet "${operation.sheetId}" was not found.`);
        }

        const deletedSheet = nextState.sheets[deleteIndex];
        nextState.sheets.splice(deleteIndex, 1);
        nextState.tables = nextState.tables.filter(
          (table) => table.range.sheetId !== deletedSheet.id,
        );

        if (nextState.activeSheetId === deletedSheet.id) {
          const nextActiveSheet =
            (operation.nextActiveSheetId
              ? nextState.sheets.find((sheet) => sheet.id === operation.nextActiveSheetId)
              : undefined) ?? nextState.sheets[Math.max(0, deleteIndex - 1)];

          nextState.activeSheetId = nextActiveSheet.id;
        }

        changed = true;
        break;
      }

      case "deleteChart": {
        const chartId = normalizeRequiredChartId(operation.chartId);
        const deleteIndex = findChartIndex(nextState, chartId);

        if (deleteIndex < 0) {
          throw new Error(`Chart "${chartId}" was not found.`);
        }

        nextState.charts = nextState.charts.filter((_chart, index) => index !== deleteIndex);
        changed = true;
        break;
      }

      case "deleteTable": {
        const tableId = normalizeRequiredTableId(operation.tableId);
        const deleteIndex = findTableIndex(nextState, tableId);

        if (deleteIndex < 0) {
          throw new Error(`Table "${tableId}" was not found.`);
        }

        nextState.tables = nextState.tables.filter((_table, index) => index !== deleteIndex);
        changed = true;
        break;
      }

      case "insertColumns": {
        assertPositiveCount(operation.count, "Column insert count");

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const insertAt = clampToRange(operation.columnIndex, 0, getSheetColumnCount(sheet));

        for (const row of sheet.cells) {
          row.splice(insertAt, 0, ...Array(operation.count).fill(""));
        }

        sheet.cellStyles = insertColumnStyles(sheet.cellStyles, insertAt, operation.count);
        sheet.columnWidths = insertColumnWidths(sheet.columnWidths, insertAt, operation.count);
        nextState.tables = nextState.tables.map((table) =>
          adjustWorkbookTableForInsertedColumns(table, sheet.id, insertAt, operation.count),
        );

        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartLayoutForInsertedColumns(
            adjustWorkbookChartForInsertedColumns(chart, sheet.id, insertAt, operation.count),
            sheet.id,
            insertAt,
            operation.count,
          ),
        );
        changed = true;
        break;
      }

      case "insertRows": {
        assertPositiveCount(operation.count, "Row insert count");

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const insertAt = clampToRange(operation.rowIndex, 0, getSheetRowCount(sheet));
        const columnCount = getSheetColumnCount(sheet);

        sheet.cells.splice(
          insertAt,
          0,
          ...Array.from({ length: operation.count }, () => Array(columnCount).fill("")),
        );
        sheet.cellStyles = insertRowStyles(sheet.cellStyles, insertAt, operation.count);
        nextState.tables = nextState.tables.map((table) =>
          adjustWorkbookTableForInsertedRows(table, sheet.id, insertAt, operation.count),
        );
        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartLayoutForInsertedRows(
            adjustWorkbookChartForInsertedRows(chart, sheet.id, insertAt, operation.count),
            sheet.id,
            insertAt,
            operation.count,
          ),
        );
        changed = true;
        break;
      }

      case "renameSheet": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const nextName = assertWorkbookSheetNameAvailable(nextState, operation.name, sheet.id);

        if (sheet.name === nextName) {
          break;
        }

        sheet.name = nextName;
        changed = true;
        break;
      }

      case "renameChart": {
        const chartId = normalizeRequiredChartId(operation.chartId);
        const chartIndex = findChartIndex(nextState, chartId);

        if (chartIndex < 0) {
          throw new Error(`Chart "${chartId}" was not found.`);
        }

        const nextName = operation.name.trim();

        if (nextName.length === 0 || nextState.charts[chartIndex].name === nextName) {
          break;
        }

        nextState.charts = nextState.charts.map((chart, index) =>
          index === chartIndex
            ? {
                ...chart,
                name: nextName,
              }
            : chart,
        );
        changed = true;
        break;
      }

      case "renameTable": {
        const tableId = normalizeRequiredTableId(operation.tableId);
        const tableIndex = findTableIndex(nextState, tableId);

        if (tableIndex < 0) {
          throw new Error(`Table "${tableId}" was not found.`);
        }

        const nextName = assertWorkbookTableNameAvailable(nextState, operation.name, tableId);

        if (nextState.tables[tableIndex].name === nextName) {
          break;
        }

        nextState.tables = nextState.tables.map((table, index) =>
          index === tableIndex
            ? {
                ...table,
                name: nextName,
              }
            : table,
        );
        changed = true;
        break;
      }

      case "replaceSheet": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const nextName =
          operation.name === undefined
            ? undefined
            : assertWorkbookSheetNameAvailable(nextState, operation.name, sheet.id);
        const nextCells = normalizeSheet(operation.rows);

        if (!matricesEqual(sheet.cells, nextCells)) {
          sheet.cells = nextCells;
          nextState.charts = nextState.charts.map((chart) =>
            clampWorkbookChartLayoutToSheet(chart, sheet),
          );
          changed = true;
        }

        if (Object.keys(sheet.cellStyles).length > 0) {
          sheet.cellStyles = {};
          changed = true;
        }

        if (Object.keys(sheet.columnWidths).length > 0) {
          sheet.columnWidths = {};
          changed = true;
        }

        if (nextState.tables.some((table) => table.range.sheetId === sheet.id)) {
          nextState.tables = nextState.tables.filter((table) => table.range.sheetId !== sheet.id);
          changed = true;
        }

        if (nextName !== undefined && nextName !== sheet.name) {
          sheet.name = nextName;
          changed = true;
        }

        if (sheet.sourceFilePath !== operation.sourceFilePath) {
          sheet.sourceFilePath = operation.sourceFilePath;
          changed = true;
        }

        break;
      }

      case "replaceSheetFromCsv": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const nextName =
          operation.name === undefined
            ? undefined
            : assertWorkbookSheetNameAvailable(nextState, operation.name, sheet.id);
        const nextCells = parseCsv(operation.content);

        if (!matricesEqual(sheet.cells, nextCells)) {
          sheet.cells = nextCells;
          nextState.charts = nextState.charts.map((chart) =>
            clampWorkbookChartLayoutToSheet(chart, sheet),
          );
          changed = true;
        }

        if (Object.keys(sheet.cellStyles).length > 0) {
          sheet.cellStyles = {};
          changed = true;
        }

        if (Object.keys(sheet.columnWidths).length > 0) {
          sheet.columnWidths = {};
          changed = true;
        }

        if (nextState.tables.some((table) => table.range.sheetId === sheet.id)) {
          nextState.tables = nextState.tables.filter((table) => table.range.sheetId !== sheet.id);
          changed = true;
        }

        if (nextName !== undefined && nextName !== sheet.name) {
          sheet.name = nextName;
          changed = true;
        }

        if (sheet.sourceFilePath !== operation.sourceFilePath) {
          sheet.sourceFilePath = operation.sourceFilePath;
          changed = true;
        }
        break;
      }

      case "resizeSheet": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const targetRowCount = Math.max(1, Math.floor(operation.rowCount));
        const targetColumnCount = Math.max(1, Math.floor(operation.columnCount));

        if (
          targetRowCount === getSheetRowCount(sheet) &&
          targetColumnCount === getSheetColumnCount(sheet)
        ) {
          break;
        }

        sheet.cells = resizeMatrix(sheet.cells, targetRowCount, targetColumnCount);
        sheet.cellStyles = filterCellStylesInBounds(
          sheet.cellStyles,
          targetRowCount,
          targetColumnCount,
        );
        sheet.columnWidths = filterColumnWidthsInBounds(sheet.columnWidths, targetColumnCount);
        nextState.tables = nextState.tables.flatMap((table) =>
          clampWorkbookTableToSheet(table, sheet),
        );
        nextState.charts = nextState.charts.map((chart) =>
          clampWorkbookChartLayoutToSheet(chart, sheet),
        );
        changed = true;
        break;
      }

      case "resizeTable": {
        const tableId = normalizeRequiredTableId(operation.tableId);
        const tableIndex = findTableIndex(nextState, tableId);

        if (tableIndex < 0) {
          throw new Error(`Table "${tableId}" was not found.`);
        }

        const currentTable = nextState.tables[tableIndex];
        const range = normalizeWorkbookTableRange(nextState, {
          ...operation.range,
          sheetId: operation.range.sheetId ?? currentTable.range.sheetId,
        });
        const nextTable: WorkbookTable = {
          ...currentTable,
          range,
          sortState: undefined,
        };

        assertWorkbookTableRangeAvailable(nextState, nextTable, tableId);

        if (workbookTablesEqual(currentTable, nextTable)) {
          break;
        }

        nextState.tables = nextState.tables.map((table, index) =>
          index === tableIndex ? nextTable : table,
        );
        changed = true;
        break;
      }

      case "setActiveSheet": {
        if (findSheetIndex(nextState, operation.sheetId) < 0) {
          throw new Error(`Sheet "${operation.sheetId}" was not found.`);
        }

        if (nextState.activeSheetId === operation.sheetId) {
          break;
        }

        nextState.activeSheetId = operation.sheetId;
        changed = true;
        break;
      }

      case "setColumnWidth": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);

        assertColumnIndex(operation.columnIndex, getSheetColumnCount(sheet), "Column width index");

        if (setColumnWidth(sheet, operation.columnIndex, operation.width)) {
          changed = true;
        }

        break;
      }

      case "setChartSpec": {
        const chartId = normalizeRequiredChartId(operation.chartId);
        const chartIndex = findChartIndex(nextState, chartId);

        if (chartIndex < 0) {
          throw new Error(`Chart "${chartId}" was not found.`);
        }

        const currentChart = nextState.charts[chartIndex];
        const nextChartCandidate: WorkbookChart = {
          ...currentChart,
          spec: cloneWorkbookChartSpec(operation.spec),
        };

        assertCreatableWorkbookChart(nextChartCandidate, nextState, "updated");

        const nextChart = clampWorkbookChartLayoutToSheet(
          nextChartCandidate,
          getSheetById(nextState, nextChartCandidate.sheetId),
        );

        assertWorkbookChartLayoutInBounds(nextChart, nextState, "updated");

        if (workbookChartsEqual(currentChart, nextChart)) {
          break;
        }

        nextState.charts = nextState.charts.map((chart, index) =>
          index === chartIndex ? nextChart : chart,
        );
        changed = true;
        break;
      }

      case "sortTable": {
        const tableId = normalizeRequiredTableId(operation.tableId);
        const tableIndex = findTableIndex(nextState, tableId);

        if (tableIndex < 0) {
          throw new Error(`Table "${tableId}" was not found.`);
        }

        const table = nextState.tables[tableIndex];
        const sortState = normalizeWorkbookTableSortState(
          table,
          operation.keys,
          operation.valueMode ?? "raw",
        );
        const sheet = getMutableSheet(nextState, clonedSheetIds, table.range.sheetId);

        if (sortWorkbookTableRows(sheet, table, sortState, operation.bodyRowOrder)) {
          nextState.tables = nextState.tables.map((entry, index) =>
            index === tableIndex
              ? {
                  ...entry,
                  sortState,
                }
              : entry,
          );
          changed = true;
          break;
        }

        if (!workbookTableSortStatesEqual(table.sortState, sortState)) {
          nextState.tables = nextState.tables.map((entry, index) =>
            index === tableIndex
              ? {
                  ...entry,
                  sortState,
                }
              : entry,
          );
          changed = true;
        }

        break;
      }

      case "setChartLayout": {
        const chartId = normalizeRequiredChartId(operation.chartId);
        const chartIndex = findChartIndex(nextState, chartId);

        if (chartIndex < 0) {
          throw new Error(`Chart "${chartId}" was not found.`);
        }

        const currentChart = nextState.charts[chartIndex];
        const nextChart: WorkbookChart = {
          ...currentChart,
          layout: normalizeWorkbookChartLayout(operation.layout),
        };

        assertWorkbookChartLayoutInBounds(nextChart, nextState, "updated");

        if (workbookChartsEqual(currentChart, nextChart)) {
          break;
        }

        nextState.charts = nextState.charts.map((chart, index) =>
          index === chartIndex ? nextChart : chart,
        );
        changed = true;
        break;
      }

      case "setSheetSourceFile": {
        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);

        if (sheet.sourceFilePath === operation.sourceFilePath) {
          break;
        }

        sheet.sourceFilePath = operation.sourceFilePath;
        changed = true;
        break;
      }

      case "setCell": {
        assertNonNegativeIndex(operation.rowIndex, "Row index");
        assertNonNegativeIndex(operation.columnIndex, "Column index");

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        ensureSheetSize(sheet, operation.rowIndex + 1, operation.columnIndex + 1);

        if (sheet.cells[operation.rowIndex][operation.columnIndex] === operation.value) {
          break;
        }

        sheet.cells[operation.rowIndex][operation.columnIndex] = operation.value;
        clearWorkbookTableSortStatesInRange(nextState, {
          columnCount: 1,
          rowCount: 1,
          sheetId: sheet.id,
          startColumn: operation.columnIndex,
          startRow: operation.rowIndex,
        });
        changed = true;
        break;
      }

      case "setCellStyle": {
        assertNonNegativeIndex(operation.rowIndex, "Row index");
        assertNonNegativeIndex(operation.columnIndex, "Column index");

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const normalizedStyle = normalizeWorkbookCellStyle(operation.style);

        ensureSheetSize(sheet, operation.rowIndex + 1, operation.columnIndex + 1);

        if (setCellStyle(sheet, operation.rowIndex, operation.columnIndex, normalizedStyle)) {
          changed = true;
        }

        break;
      }

      case "setRange": {
        if (operation.values.length === 0) {
          break;
        }

        assertNonNegativeIndex(operation.startRow, "Start row");
        assertNonNegativeIndex(operation.startColumn, "Start column");

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const maxColumnCount = Math.max(0, ...operation.values.map((row) => row.length));
        let rangeChanged = false;

        if (maxColumnCount === 0) {
          break;
        }

        ensureSheetSize(
          sheet,
          operation.startRow + operation.values.length,
          operation.startColumn + maxColumnCount,
        );

        for (let rowOffset = 0; rowOffset < operation.values.length; rowOffset += 1) {
          const sourceRow = operation.values[rowOffset];
          const targetRow = sheet.cells[operation.startRow + rowOffset];

          for (let columnOffset = 0; columnOffset < sourceRow.length; columnOffset += 1) {
            const nextValue = sourceRow[columnOffset] ?? "";
            const targetColumn = operation.startColumn + columnOffset;

            if (targetRow[targetColumn] === nextValue) {
              continue;
            }

            targetRow[targetColumn] = nextValue;
            rangeChanged = true;
            changed = true;
          }
        }

        if (rangeChanged) {
          clearWorkbookTableSortStatesInRange(nextState, {
            columnCount: maxColumnCount,
            rowCount: operation.values.length,
            sheetId: sheet.id,
            startColumn: operation.startColumn,
            startRow: operation.startRow,
          });
        }

        break;
      }

      case "setRangeStyle": {
        assertNonNegativeIndex(operation.startRow, "Start row");
        assertNonNegativeIndex(operation.startColumn, "Start column");

        const rowCount = Math.max(0, Math.floor(operation.rowCount));
        const columnCount = Math.max(0, Math.floor(operation.columnCount));

        if (rowCount === 0 || columnCount === 0) {
          break;
        }

        const sheet = getMutableSheet(nextState, clonedSheetIds, operation.sheetId);
        const normalizedStyle = normalizeWorkbookCellStyle(operation.style);

        ensureSheetSize(sheet, operation.startRow + rowCount, operation.startColumn + columnCount);

        for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
          for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
            if (
              setCellStyle(
                sheet,
                operation.startRow + rowOffset,
                operation.startColumn + columnOffset,
                normalizedStyle,
              )
            ) {
              changed = true;
            }
          }
        }

        break;
      }
    }
  }

  if (!changed) {
    return {
      changed: false,
      state: previousState,
    };
  }

  if (!request.dryRun) {
    nextState.version = previousState.version + 1;
  }

  return {
    changed: true,
    state: nextState,
  };
}

function assertPositiveCount(value: number, label: string) {
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`${label} must be a positive integer.`);
  }
}

function assertNonNegativeIndex(value: number, label: string) {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`${label} must be a non-negative integer.`);
  }
}

function assertColumnIndex(value: number, columnCount: number, label: string) {
  assertNonNegativeIndex(value, label);

  if (value >= columnCount) {
    throw new Error(`${label} must be inside the sheet column bounds.`);
  }
}

function assertNonNegativeInteger(value: number, label: string) {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`${label} must be a non-negative integer.`);
  }
}

function normalizePositiveInteger(value: number, label: string): number {
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`${label} must be a positive integer.`);
  }

  return value;
}

function normalizeNonNegativeInteger(value: number, label: string): number {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`${label} must be a non-negative integer.`);
  }

  return value;
}

function assertNonNegativeFiniteNumber(value: number, label: string) {
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`${label} must be a non-negative finite number.`);
  }
}

function assertMinimumFiniteNumber(value: number, minimum: number, label: string) {
  if (!Number.isFinite(value) || value < minimum) {
    throw new Error(`${label} must be at least ${minimum}.`);
  }
}

function assertMaximumFiniteNumber(value: number, maximum: number, label: string) {
  if (!Number.isFinite(value) || value > maximum) {
    throw new Error(`${label} must be at most ${maximum}.`);
  }
}

function clampToRange(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(Math.floor(value), max));
}

function cloneWorkbookChartSpec(spec: WorkbookChartSpec): WorkbookChartSpec {
  return spec.family === "cartesian"
    ? {
        ...spec,
        source: {
          ...spec.source,
          range: {
            ...spec.source.range,
          },
        },
        valueDimensions: [...spec.valueDimensions],
      }
    : {
        ...spec,
        source: {
          ...spec.source,
          range: {
            ...spec.source.range,
          },
        },
      };
}

function cloneWorkbookChartLayout(layout: WorkbookChartLayout): WorkbookChartLayout {
  return {
    height: layout.height,
    offsetX: layout.offsetX,
    offsetY: layout.offsetY,
    startColumn: layout.startColumn,
    startRow: layout.startRow,
    width: layout.width,
    zIndex: layout.zIndex,
  };
}

function cloneWorkbookTableSortState(sortState: WorkbookTableSortState): WorkbookTableSortState {
  return {
    keys: sortState.keys.map((key) => ({ ...key })),
    valueMode: sortState.valueMode,
  };
}

function cloneRangeValues(values: readonly (readonly string[])[]): string[][] {
  return values.map((row) => [...row]);
}

function createWorkbookRangeFromSheetRange(range: SheetRangeResult): WorkbookTableRange {
  return {
    columnCount: range.columnCount,
    rowCount: range.rowCount,
    sheetId: range.sheetId,
    startColumn: range.startColumn,
    startRow: range.startRow,
  };
}

function createWorkbookRangeFromClipboardTable(
  request: Pick<PasteRangeRequest, "startColumn" | "startRow">,
  sheetId: string,
  table: ClipboardTable,
): WorkbookTableRange {
  return {
    columnCount: table.range.columnCount,
    rowCount: table.range.rowCount,
    sheetId,
    startColumn: request.startColumn + table.range.startColumnOffset,
    startRow: request.startRow + table.range.startRowOffset,
  };
}

function workbookTableRangeContains(outer: WorkbookTableRange, inner: WorkbookTableRange): boolean {
  return (
    outer.sheetId === inner.sheetId &&
    inner.startRow >= outer.startRow &&
    inner.startColumn >= outer.startColumn &&
    inner.startRow + inner.rowCount <= outer.startRow + outer.rowCount &&
    inner.startColumn + inner.columnCount <= outer.startColumn + outer.columnCount
  );
}

function workbookRangeIntersectsTableHeader(
  range: WorkbookTableRange,
  table: WorkbookTable,
): boolean {
  return (
    range.sheetId === table.range.sheetId &&
    range.startRow <= table.range.startRow &&
    range.startRow + range.rowCount > table.range.startRow &&
    range.startColumn < table.range.startColumn + table.range.columnCount &&
    range.startColumn + range.columnCount > table.range.startColumn
  );
}

function getWorkbookRangeRemainders(
  previousRange: WorkbookTableRange,
  nextRange: WorkbookTableRange,
): WorkbookTableRange[] {
  const remainders: WorkbookTableRange[] = [];
  const previousEndRow = previousRange.startRow + previousRange.rowCount;
  const previousEndColumn = previousRange.startColumn + previousRange.columnCount;
  const nextEndRow = nextRange.startRow + nextRange.rowCount;
  const nextEndColumn = nextRange.startColumn + nextRange.columnCount;

  if (nextEndRow < previousEndRow) {
    remainders.push({
      columnCount: previousRange.columnCount,
      rowCount: previousEndRow - nextEndRow,
      sheetId: previousRange.sheetId,
      startColumn: previousRange.startColumn,
      startRow: nextEndRow,
    });
  }

  if (nextEndColumn < previousEndColumn) {
    remainders.push({
      columnCount: previousEndColumn - nextEndColumn,
      rowCount: Math.min(previousEndRow, nextEndRow) - previousRange.startRow,
      sheetId: previousRange.sheetId,
      startColumn: nextEndColumn,
      startRow: previousRange.startRow,
    });
  }

  return remainders.filter((range) => range.rowCount > 0 && range.columnCount > 0);
}

function getAvailablePastedWorkbookTableName(usedNameKeys: Set<string>, name: string): string {
  const baseName = name.trim() || "Table";
  const baseNameKey = getWorkbookTableNameKey(baseName);

  if (!usedNameKeys.has(baseNameKey)) {
    usedNameKeys.add(baseNameKey);
    return baseName;
  }

  for (let index = 1; ; index += 1) {
    const candidate = index === 1 ? `${baseName} Copy` : `${baseName} Copy ${index}`;
    const candidateKey = getWorkbookTableNameKey(candidate);

    if (!usedNameKeys.has(candidateKey)) {
      usedNameKeys.add(candidateKey);
      return candidate;
    }
  }
}

function createWorkbookChart(
  id: string,
  name: string,
  spec: WorkbookChartSpec,
  layout: WorkbookChartLayout | undefined,
  sheetId: string,
  workbook: Pick<WorkbookState, "charts" | "sheets">,
): WorkbookChart {
  const nextSpec = cloneWorkbookChartSpec(spec);

  return {
    id,
    layout: layout
      ? normalizeWorkbookChartLayout(layout)
      : createDefaultWorkbookChartLayout(nextSpec, sheetId, workbook),
    name,
    sheetId,
    spec: nextSpec,
  };
}

function createDefaultWorkbookChartLayout(
  spec: WorkbookChartSpec,
  sheetId: string,
  workbook: Pick<WorkbookState, "charts" | "sheets">,
): WorkbookChartLayout {
  const sheet = getSheetById(workbook, sheetId);
  const columnCount = getSheetColumnCount(sheet);
  const rowCount = getSheetRowCount(sheet);
  const preferredColumn = spec.source.range.startColumn + spec.source.range.columnCount + 1;

  return {
    height: DEFAULT_CHART_LAYOUT_HEIGHT,
    offsetX: 0,
    offsetY: 0,
    startColumn: clampToRange(preferredColumn, 0, columnCount - 1),
    startRow: clampToRange(spec.source.range.startRow, 0, rowCount - 1),
    width: DEFAULT_CHART_LAYOUT_WIDTH,
    zIndex: getNextWorkbookChartZIndex(workbook),
  };
}

function getNextWorkbookChartZIndex(workbook: Pick<WorkbookState, "charts">): number {
  return Math.max(-1, ...workbook.charts.map((chart) => Math.floor(chart.layout.zIndex))) + 1;
}

function createChartId(nextChartNumber: number): string {
  return `chart-${nextChartNumber}`;
}

function findChartIndex(workbook: WorkbookState, chartId: string): number {
  return workbook.charts.findIndex((chart) => chart.id === chartId);
}

function createTableId(nextTableNumber: number): string {
  return `table-${nextTableNumber}`;
}

function getNextWorkbookTableNumberForId(tableId: string): number {
  const match = /^table-(\d+)$/u.exec(tableId);

  return match ? Number.parseInt(match[1], 10) + 1 : 1;
}

function findTableIndex(workbook: Pick<WorkbookState, "tables">, tableId: string): number {
  return workbook.tables.findIndex((table) => table.id === tableId);
}

function normalizeOptionalChartId(chartId?: string): string | undefined {
  const nextChartId = chartId?.trim();

  return nextChartId && nextChartId.length > 0 ? nextChartId : undefined;
}

function normalizeOptionalTableId(tableId?: string): string | undefined {
  const nextTableId = tableId?.trim();

  return nextTableId && nextTableId.length > 0 ? nextTableId : undefined;
}

function normalizeRequiredChartId(chartId: string): string {
  const nextChartId = chartId.trim();

  if (nextChartId.length === 0) {
    throw new Error("Chart id must be a non-empty string.");
  }

  return nextChartId;
}

function normalizeRequiredTableId(tableId: string): string {
  const nextTableId = tableId.trim();

  if (nextTableId.length === 0) {
    throw new Error("Table id must be a non-empty string.");
  }

  return nextTableId;
}

function getWorkbookChartSheetReferences(
  workbook: Pick<WorkbookState, "sheets">,
): WorkbookChartSheetReference[] {
  return workbook.sheets.map((sheet) => ({
    columnCount: getSheetColumnCount(sheet),
    id: sheet.id,
    rowCount: getSheetRowCount(sheet),
  }));
}

function createChartRangeFromRequest(
  sheet: WorkbookSheet,
  sourceRange?: CreateChartSourceRange,
): WorkbookChartRange {
  if (!sourceRange) {
    const usedRange = getUsedRange(sheet);

    if (usedRange.rowCount === 0 || usedRange.columnCount === 0) {
      throw new Error("Chart source range was omitted, but the target sheet has no used cells.");
    }

    return {
      columnCount: usedRange.columnCount,
      rowCount: usedRange.rowCount,
      sheetId: sheet.id,
      startColumn: usedRange.startColumn,
      startRow: usedRange.startRow,
    };
  }

  assertNonNegativeInteger(sourceRange.startRow, "Chart source start row");
  assertNonNegativeInteger(sourceRange.startColumn, "Chart source start column");
  assertPositiveCount(sourceRange.rowCount, "Chart source row count");
  assertPositiveCount(sourceRange.columnCount, "Chart source column count");

  return {
    columnCount: sourceRange.columnCount,
    rowCount: sourceRange.rowCount,
    sheetId: sheet.id,
    startColumn: sourceRange.startColumn,
    startRow: sourceRange.startRow,
  };
}

function getDefaultSecondaryChartDimension(
  primaryDimension: number,
  dimensionCount: number,
): number {
  return (
    Array.from({ length: dimensionCount }, (_value, index) => index).find(
      (dimension) => dimension !== primaryDimension,
    ) ?? 1
  );
}

function getDefaultValueChartDimensions(
  categoryDimension: number,
  dimensionCount: number,
): number[] {
  const valueDimensions = Array.from({ length: dimensionCount }, (_value, index) => index).filter(
    (dimension) => dimension !== categoryDimension,
  );

  return valueDimensions.length > 0 ? valueDimensions : [1];
}

function normalizeFormatCellsRange(
  workbook: WorkbookState,
  range: SheetRangeRequest,
): Required<SheetRangeRequest> {
  const sheet = getWorkbookSheet(workbook, range.sheetId);

  assertNonNegativeInteger(range.startRow, "Format range start row");
  assertNonNegativeInteger(range.startColumn, "Format range start column");
  assertPositiveCount(range.rowCount, "Format range row count");
  assertPositiveCount(range.columnCount, "Format range column count");

  return {
    columnCount: range.columnCount,
    rowCount: range.rowCount,
    sheetId: sheet.id,
    startColumn: range.startColumn,
    startRow: range.startRow,
  };
}

function assertCreatableWorkbookChart(
  chart: WorkbookChart,
  workbook: Pick<WorkbookState, "sheets">,
  verb: "added" | "updated",
) {
  const issues = getWorkbookChartValidationIssues(chart, getWorkbookChartSheetReferences(workbook));

  if (issues.length === 0) {
    return;
  }

  throw new Error(
    `Chart "${chart.id}" cannot be ${verb}: ${issues.map((issue) => issue.message).join(" ")}`,
  );
}

function workbookChartsEqual(left: WorkbookChart, right: WorkbookChart): boolean {
  return (
    left.id === right.id &&
    workbookChartLayoutsEqual(left.layout, right.layout) &&
    left.name === right.name &&
    left.sheetId === right.sheetId &&
    workbookChartSpecsEqual(left.spec, right.spec)
  );
}

function workbookTablesEqual(left: WorkbookTable, right: WorkbookTable): boolean {
  return (
    left.hasHeaderRow === right.hasHeaderRow &&
    left.id === right.id &&
    left.name === right.name &&
    workbookTableRangesEqual(left.range, right.range) &&
    workbookTableSortStatesEqual(left.sortState, right.sortState)
  );
}

function workbookTableRangesEqual(left: WorkbookTableRange, right: WorkbookTableRange): boolean {
  return (
    left.sheetId === right.sheetId &&
    left.startRow === right.startRow &&
    left.startColumn === right.startColumn &&
    left.rowCount === right.rowCount &&
    left.columnCount === right.columnCount
  );
}

function workbookTableSortStatesEqual(
  left: WorkbookTableSortState | undefined,
  right: WorkbookTableSortState | undefined,
): boolean {
  if (!left || !right) {
    return left === right;
  }

  return (
    left.valueMode === right.valueMode &&
    left.keys.length === right.keys.length &&
    left.keys.every(
      (key, index) =>
        key.columnIndex === right.keys[index]?.columnIndex &&
        key.direction === right.keys[index]?.direction,
    )
  );
}

function workbookChartSpecsEqual(left: WorkbookChartSpec, right: WorkbookChartSpec): boolean {
  if (
    left.family !== right.family ||
    left.chartType !== right.chartType ||
    left.source.seriesLayoutBy !== right.source.seriesLayoutBy ||
    left.source.sourceHeader !== right.source.sourceHeader ||
    !workbookChartRangesEqual(left.source.range, right.source.range)
  ) {
    return false;
  }

  if (left.family === "cartesian" && right.family === "cartesian") {
    return (
      left.categoryDimension === right.categoryDimension &&
      left.smooth === right.smooth &&
      left.stacked === right.stacked &&
      arraysEqual(left.valueDimensions, right.valueDimensions)
    );
  }

  if (left.family === "pie" && right.family === "pie") {
    return (
      left.nameDimension === right.nameDimension && left.valueDimension === right.valueDimension
    );
  }

  return false;
}

function workbookChartRangesEqual(left: WorkbookChartRange, right: WorkbookChartRange): boolean {
  return (
    left.sheetId === right.sheetId &&
    left.startRow === right.startRow &&
    left.startColumn === right.startColumn &&
    left.rowCount === right.rowCount &&
    left.columnCount === right.columnCount
  );
}

function workbookChartLayoutsEqual(left: WorkbookChartLayout, right: WorkbookChartLayout): boolean {
  return (
    left.height === right.height &&
    left.offsetX === right.offsetX &&
    left.offsetY === right.offsetY &&
    left.startColumn === right.startColumn &&
    left.startRow === right.startRow &&
    left.width === right.width &&
    left.zIndex === right.zIndex
  );
}

function arraysEqual<T>(left: readonly T[], right: readonly T[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function normalizeWorkbookChartLayout(layout: WorkbookChartLayout): WorkbookChartLayout {
  const nextLayout = {
    height: layout.height,
    offsetX: layout.offsetX,
    offsetY: layout.offsetY,
    startColumn: layout.startColumn,
    startRow: layout.startRow,
    width: layout.width,
    zIndex: layout.zIndex,
  };

  assertNonNegativeInteger(nextLayout.startRow, "Chart layout start row");
  assertNonNegativeInteger(nextLayout.startColumn, "Chart layout start column");
  assertNonNegativeFiniteNumber(nextLayout.offsetX, "Chart layout offset x");
  assertNonNegativeFiniteNumber(nextLayout.offsetY, "Chart layout offset y");
  assertMinimumFiniteNumber(nextLayout.width, MIN_CHART_LAYOUT_WIDTH, "Chart layout width");
  assertMinimumFiniteNumber(nextLayout.height, MIN_CHART_LAYOUT_HEIGHT, "Chart layout height");
  assertNonNegativeInteger(nextLayout.zIndex, "Chart layout z-index");

  return nextLayout;
}

function assertWorkbookChartLayoutInBounds(
  chart: WorkbookChart,
  workbook: Pick<WorkbookState, "sheets">,
  verb: "added" | "updated",
) {
  const sheet = getSheetById(workbook, chart.sheetId);
  const rowCount = getSheetRowCount(sheet);
  const columnCount = getSheetColumnCount(sheet);

  if (chart.layout.startRow >= rowCount || chart.layout.startColumn >= columnCount) {
    throw new Error(
      `Chart "${chart.id}" cannot be ${verb}: chart layout anchor must be inside the chart sheet bounds.`,
    );
  }
}

function clampWorkbookChartLayoutToSheet(
  chart: WorkbookChart,
  sheet: WorkbookSheet,
): WorkbookChart {
  if (chart.sheetId !== sheet.id) {
    return chart;
  }

  const layout = chart.layout;
  const nextLayout = {
    ...layout,
    startColumn: clampToRange(layout.startColumn, 0, getSheetColumnCount(sheet) - 1),
    startRow: clampToRange(layout.startRow, 0, getSheetRowCount(sheet) - 1),
  };

  return workbookChartLayoutsEqual(layout, nextLayout)
    ? chart
    : {
        ...chart,
        layout: nextLayout,
      };
}

function adjustWorkbookChartLayoutForInsertedRows(
  chart: WorkbookChart,
  sheetId: string,
  rowIndex: number,
  count: number,
): WorkbookChart {
  if (chart.sheetId !== sheetId || rowIndex > chart.layout.startRow) {
    return chart;
  }

  return {
    ...chart,
    layout: {
      ...chart.layout,
      startRow: chart.layout.startRow + count,
    },
  };
}

function adjustWorkbookChartLayoutForInsertedColumns(
  chart: WorkbookChart,
  sheetId: string,
  columnIndex: number,
  count: number,
): WorkbookChart {
  if (chart.sheetId !== sheetId || columnIndex > chart.layout.startColumn) {
    return chart;
  }

  return {
    ...chart,
    layout: {
      ...chart.layout,
      startColumn: chart.layout.startColumn + count,
    },
  };
}

function adjustWorkbookChartLayoutForDeletedRows(
  chart: WorkbookChart,
  sheetId: string,
  rowIndex: number,
  count: number,
  nextRowCount: number,
): WorkbookChart {
  if (chart.sheetId !== sheetId) {
    return chart;
  }

  const nextStartRow = clampToRange(
    adjustPointForDeletion(chart.layout.startRow, rowIndex, count),
    0,
    nextRowCount - 1,
  );

  return nextStartRow === chart.layout.startRow
    ? chart
    : {
        ...chart,
        layout: {
          ...chart.layout,
          startRow: nextStartRow,
        },
      };
}

function adjustWorkbookChartLayoutForDeletedColumns(
  chart: WorkbookChart,
  sheetId: string,
  columnIndex: number,
  count: number,
  nextColumnCount: number,
): WorkbookChart {
  if (chart.sheetId !== sheetId) {
    return chart;
  }

  const nextStartColumn = clampToRange(
    adjustPointForDeletion(chart.layout.startColumn, columnIndex, count),
    0,
    nextColumnCount - 1,
  );

  return nextStartColumn === chart.layout.startColumn
    ? chart
    : {
        ...chart,
        layout: {
          ...chart.layout,
          startColumn: nextStartColumn,
        },
      };
}

function adjustPointForDeletion(point: number, deleteIndex: number, count: number): number {
  if (point < deleteIndex) {
    return point;
  }

  if (point >= deleteIndex + count) {
    return point - count;
  }

  return deleteIndex;
}

function updateWorkbookChartRange(chart: WorkbookChart, range: WorkbookChartRange): WorkbookChart {
  const currentRange = chart.spec.source.range;

  if (
    currentRange.sheetId === range.sheetId &&
    currentRange.startRow === range.startRow &&
    currentRange.startColumn === range.startColumn &&
    currentRange.rowCount === range.rowCount &&
    currentRange.columnCount === range.columnCount
  ) {
    return chart;
  }

  return {
    ...chart,
    spec: {
      ...chart.spec,
      source: {
        ...chart.spec.source,
        range,
      },
    },
  };
}

function updateWorkbookTableRange(table: WorkbookTable, range: WorkbookTableRange): WorkbookTable {
  if (
    table.range.sheetId === range.sheetId &&
    table.range.startRow === range.startRow &&
    table.range.startColumn === range.startColumn &&
    table.range.rowCount === range.rowCount &&
    table.range.columnCount === range.columnCount
  ) {
    return table;
  }

  return {
    ...table,
    range,
    sortState: undefined,
  };
}

function clampWorkbookTableToSheet(table: WorkbookTable, sheet: WorkbookSheet): WorkbookTable[] {
  if (table.range.sheetId !== sheet.id) {
    return [table];
  }

  const rowCount = getSheetRowCount(sheet);
  const columnCount = getSheetColumnCount(sheet);
  const startRow = Math.min(table.range.startRow, rowCount);
  const startColumn = Math.min(table.range.startColumn, columnCount);
  const nextRowCount = Math.max(0, Math.min(table.range.rowCount, rowCount - startRow));
  const nextColumnCount = Math.max(0, Math.min(table.range.columnCount, columnCount - startColumn));

  if (nextRowCount === 0 || nextColumnCount === 0) {
    return [];
  }

  return [
    updateWorkbookTableRange(table, {
      columnCount: nextColumnCount,
      rowCount: nextRowCount,
      sheetId: sheet.id,
      startColumn,
      startRow,
    }),
  ];
}

function adjustRangeAxisForDeletion(
  start: number,
  count: number,
  deleteStart: number,
  deleteCount: number,
): { start: number; count: number } {
  const end = start + count;
  const deleteEnd = deleteStart + deleteCount;

  if (deleteEnd <= start) {
    return {
      count,
      start: start - deleteCount,
    };
  }

  if (deleteStart >= end) {
    return {
      count,
      start,
    };
  }

  const survivorCountBeforeDelete = Math.max(0, Math.min(end, deleteStart) - start);
  const survivorCountAfterDelete = Math.max(0, end - Math.max(start, deleteEnd));
  const nextCount = survivorCountBeforeDelete + survivorCountAfterDelete;

  if (nextCount === 0) {
    return {
      count: 0,
      start: Math.max(0, Math.min(start, deleteStart)),
    };
  }

  if (survivorCountBeforeDelete === 0) {
    return {
      count: nextCount,
      start: Math.max(0, Math.min(start, deleteStart)),
    };
  }

  return {
    count: nextCount,
    start,
  };
}

function normalizeWorkbookTableSortState(
  table: WorkbookTable,
  keys: WorkbookTableSortKey[],
  valueMode: WorkbookTableSortValueMode,
): WorkbookTableSortState {
  if (keys.length === 0) {
    throw new Error("Table sort must include at least one sort key.");
  }

  if (valueMode !== "raw" && valueMode !== "display") {
    throw new Error('Table sort valueMode must be "raw" or "display".');
  }

  return {
    keys: keys.map((key) => {
      const columnIndex = normalizeNonNegativeInteger(key.columnIndex, "Table sort column index");

      if (
        columnIndex < table.range.startColumn ||
        columnIndex >= table.range.startColumn + table.range.columnCount
      ) {
        throw new Error(`Table sort column ${columnIndex} is outside table "${table.id}".`);
      }

      if (key.direction !== "ascending" && key.direction !== "descending") {
        throw new Error('Table sort direction must be "ascending" or "descending".');
      }

      return {
        columnIndex,
        direction: key.direction,
      };
    }),
    valueMode,
  };
}

function sortWorkbookTableRows(
  sheet: WorkbookSheet,
  table: WorkbookTable,
  sortState: WorkbookTableSortState,
  bodyRowOrder?: number[],
): boolean {
  const headerOffset = table.hasHeaderRow ? 1 : 0;
  const bodyStartRow = table.range.startRow + headerOffset;
  const bodyRowCount = table.range.rowCount - headerOffset;

  if (bodyRowCount <= 1) {
    return false;
  }

  const sortedRows =
    bodyRowOrder === undefined
      ? getSortedWorkbookTableBodyRows(sheet, table, sortState)
      : normalizeWorkbookTableBodyRowOrder(bodyRowOrder, bodyStartRow, bodyRowCount);
  const originalRows = Array.from(
    { length: bodyRowCount },
    (_value, index) => bodyStartRow + index,
  );

  if (sortedRows.every((rowIndex, index) => rowIndex === originalRows[index])) {
    return false;
  }

  const startColumn = table.range.startColumn;
  const endColumn = startColumn + table.range.columnCount;
  const sortedValues = sortedRows.map((sourceRowIndex) =>
    sheet.cells[sourceRowIndex].slice(startColumn, endColumn),
  );

  for (let rowOffset = 0; rowOffset < bodyRowCount; rowOffset += 1) {
    const targetRow = sheet.cells[bodyStartRow + rowOffset];
    const sourceValues = sortedValues[rowOffset];

    for (let columnOffset = 0; columnOffset < table.range.columnCount; columnOffset += 1) {
      targetRow[startColumn + columnOffset] = sourceValues[columnOffset] ?? "";
    }
  }

  sheet.cellStyles = sortWorkbookTableCellStyles(sheet.cellStyles, table, sortedRows);
  return true;
}

function getSortedWorkbookTableBodyRows(
  sheet: WorkbookSheet,
  table: WorkbookTable,
  sortState: WorkbookTableSortState,
): number[] {
  const headerOffset = table.hasHeaderRow ? 1 : 0;
  const bodyStartRow = table.range.startRow + headerOffset;
  const bodyRowCount = table.range.rowCount - headerOffset;

  return Array.from({ length: bodyRowCount }, (_value, index) => bodyStartRow + index)
    .map((rowIndex, originalIndex) => ({ originalIndex, rowIndex }))
    .sort((left, right) => {
      for (const key of sortState.keys) {
        const comparison = compareWorkbookTableSortValues(
          sheet.cells[left.rowIndex]?.[key.columnIndex] ?? "",
          sheet.cells[right.rowIndex]?.[key.columnIndex] ?? "",
        );

        if (comparison !== 0) {
          return key.direction === "ascending" ? comparison : -comparison;
        }
      }

      return left.originalIndex - right.originalIndex;
    })
    .map((entry) => entry.rowIndex);
}

function normalizeWorkbookTableBodyRowOrder(
  bodyRowOrder: number[],
  bodyStartRow: number,
  bodyRowCount: number,
): number[] {
  if (bodyRowOrder.length !== bodyRowCount) {
    throw new Error("Table body row order must include every table body row exactly once.");
  }

  const allowedRows = new Set(
    Array.from({ length: bodyRowCount }, (_value, index) => bodyStartRow + index),
  );
  const seenRows = new Set<number>();

  for (const rowIndex of bodyRowOrder) {
    if (!Number.isInteger(rowIndex) || !allowedRows.has(rowIndex) || seenRows.has(rowIndex)) {
      throw new Error("Table body row order must include every table body row exactly once.");
    }

    seenRows.add(rowIndex);
  }

  return [...bodyRowOrder];
}

export function compareWorkbookTableSortValues(left: string, right: string): number {
  const leftBlank = left.trim().length === 0;
  const rightBlank = right.trim().length === 0;

  if (leftBlank || rightBlank) {
    if (leftBlank === rightBlank) {
      return 0;
    }

    return leftBlank ? 1 : -1;
  }

  const leftNumber = Number(left);
  const rightNumber = Number(right);

  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return leftNumber - rightNumber;
  }

  return left.localeCompare(right, undefined, {
    numeric: true,
    sensitivity: "base",
  });
}

function sortWorkbookTableCellStyles(
  styles: Record<string, WorkbookCellStyle>,
  table: WorkbookTable,
  sortedRows: number[],
): Record<string, WorkbookCellStyle> {
  const headerOffset = table.hasHeaderRow ? 1 : 0;
  const bodyStartRow = table.range.startRow + headerOffset;
  const bodyEndRow = table.range.startRow + table.range.rowCount;
  const startColumn = table.range.startColumn;
  const endColumn = table.range.startColumn + table.range.columnCount;
  const nextStyles: Record<string, WorkbookCellStyle> = {};

  for (const [key, style] of Object.entries(styles)) {
    const [rowText, columnText] = key.split(":");
    const rowIndex = Number.parseInt(rowText, 10);
    const columnIndex = Number.parseInt(columnText, 10);

    if (
      rowIndex >= bodyStartRow &&
      rowIndex < bodyEndRow &&
      columnIndex >= startColumn &&
      columnIndex < endColumn
    ) {
      continue;
    }

    nextStyles[key] = cloneWorkbookCellStyle(style);
  }

  for (let rowOffset = 0; rowOffset < sortedRows.length; rowOffset += 1) {
    const targetRowIndex = bodyStartRow + rowOffset;
    const sourceRowIndex = sortedRows[rowOffset];

    for (let columnIndex = startColumn; columnIndex < endColumn; columnIndex += 1) {
      const style = styles[`${sourceRowIndex}:${columnIndex}`];

      if (style) {
        nextStyles[`${targetRowIndex}:${columnIndex}`] = cloneWorkbookCellStyle(style);
      }
    }
  }

  return nextStyles;
}

function createWorkbookSheet(
  name: string,
  rowCount: number,
  columnCount: number,
  id = createSheetId(),
): WorkbookSheet {
  registerSheetId(id);

  return {
    cells: createSheet(rowCount, columnCount),
    cellStyles: {},
    columnWidths: {},
    id,
    name,
  };
}

function assertWorkbookSheetNameAvailable(
  workbook: Pick<WorkbookState, "sheets">,
  name: string,
  allowedSheetId?: string,
): string {
  const normalizedName = normalizeWorkbookSheetName(name);
  const nameKey = getWorkbookSheetNameKey(normalizedName);
  const existingSheet = workbook.sheets.find(
    (sheet) => sheet.id !== allowedSheetId && getWorkbookSheetNameKey(sheet.name) === nameKey,
  );

  if (existingSheet) {
    throw new Error(
      `Sheet name "${normalizedName}" already exists as "${existingSheet.name}". Sheet names must be unique case-insensitively.`,
    );
  }

  return normalizedName;
}

function getNextAvailableWorkbookSheetName(
  workbook: Pick<WorkbookState, "nextSheetNumber" | "sheets">,
): {
  name: string;
  nextSheetNumber: number;
} {
  for (let sheetNumber = Math.max(1, Math.floor(workbook.nextSheetNumber)); ; sheetNumber += 1) {
    const name = `Sheet ${sheetNumber}`;

    if (
      !workbook.sheets.some(
        (sheet) => getWorkbookSheetNameKey(sheet.name) === getWorkbookSheetNameKey(name),
      )
    ) {
      return {
        name,
        nextSheetNumber: sheetNumber + 1,
      };
    }
  }
}

function getNextAvailableWorkbookTableName(
  workbook: Pick<WorkbookState, "nextTableNumber" | "tables">,
): {
  name: string;
  nextTableNumber: number;
} {
  for (let tableNumber = Math.max(1, Math.floor(workbook.nextTableNumber)); ; tableNumber += 1) {
    const name = `Table ${tableNumber}`;

    if (
      !workbook.tables.some(
        (table) => getWorkbookTableNameKey(table.name) === getWorkbookTableNameKey(name),
      )
    ) {
      return {
        name,
        nextTableNumber: tableNumber + 1,
      };
    }
  }
}

function getWorkbookSheetNameKey(name: string): string {
  return name.trim().toLowerCase();
}

function normalizeWorkbookTableName(name: string): string {
  const normalizedName = name.trim();

  if (normalizedName.length === 0) {
    throw new Error("Table name is required.");
  }

  return normalizedName;
}

function getWorkbookTableNameKey(name: string): string {
  return name.trim().toLowerCase();
}

function assertWorkbookTableNameAvailable(
  workbook: Pick<WorkbookState, "tables">,
  name: string,
  allowedTableId?: string,
): string {
  const normalizedName = normalizeWorkbookTableName(name);
  const nameKey = getWorkbookTableNameKey(normalizedName);
  const existingTable = workbook.tables.find(
    (table) => table.id !== allowedTableId && getWorkbookTableNameKey(table.name) === nameKey,
  );

  if (existingTable) {
    throw new Error(
      `Table name "${normalizedName}" already exists as "${existingTable.name}". Table names must be unique case-insensitively.`,
    );
  }

  return normalizedName;
}

function normalizeWorkbookTableRange(
  workbook: WorkbookState,
  range: Omit<WorkbookTableRange, "sheetId"> & { sheetId?: string },
): WorkbookTableRange {
  const sheet = getWorkbookSheet(workbook, range.sheetId ?? workbook.activeSheetId);
  const startRow = normalizeNonNegativeInteger(range.startRow, "Table start row");
  const startColumn = normalizeNonNegativeInteger(range.startColumn, "Table start column");
  const rowCount = normalizePositiveInteger(range.rowCount, "Table row count");
  const columnCount = normalizePositiveInteger(range.columnCount, "Table column count");

  if (startRow + rowCount > getSheetRowCount(sheet)) {
    throw new Error("Table range must fit within the sheet row bounds.");
  }

  if (startColumn + columnCount > getSheetColumnCount(sheet)) {
    throw new Error("Table range must fit within the sheet column bounds.");
  }

  return {
    columnCount,
    rowCount,
    sheetId: sheet.id,
    startColumn,
    startRow,
  };
}

function assertWorkbookTableRangeAvailable(
  workbook: Pick<WorkbookState, "tables">,
  table: WorkbookTable,
  allowedTableId?: string,
) {
  const overlappingTable = workbook.tables.find(
    (existingTable) =>
      existingTable.id !== allowedTableId &&
      existingTable.range.sheetId === table.range.sheetId &&
      rangesOverlap(existingTable.range, table.range),
  );

  if (overlappingTable) {
    throw new Error(`Table "${table.name}" overlaps table "${overlappingTable.name}".`);
  }
}

function clearWorkbookTableSortStatesInRange(
  workbook: Pick<WorkbookState, "tables">,
  range: WorkbookTableRange,
): boolean {
  let changed = false;

  const tables = workbook.tables.map((table) => {
    if (
      !table.sortState ||
      table.range.sheetId !== range.sheetId ||
      !rangesOverlap(table.range, range)
    ) {
      return table;
    }

    changed = true;
    return {
      ...table,
      sortState: undefined,
    };
  });

  if (changed) {
    workbook.tables = tables;
  }

  return changed;
}

function rangesOverlap(left: WorkbookTableRange, right: WorkbookTableRange): boolean {
  return (
    left.startRow < right.startRow + right.rowCount &&
    right.startRow < left.startRow + left.rowCount &&
    left.startColumn < right.startColumn + right.columnCount &&
    right.startColumn < left.startColumn + left.columnCount
  );
}

function createSheetId(): string {
  const sheetId = `sheet-${nextSheetIdSequence}`;

  nextSheetIdSequence += 1;
  return sheetId;
}

export function syncSheetIdSequence(workbook: WorkbookState) {
  for (const sheet of workbook.sheets) {
    registerSheetId(sheet.id);
  }
}

function ensureSheetSize(
  sheet: WorkbookSheet,
  minimumRowCount: number,
  minimumColumnCount: number,
) {
  const targetRowCount = Math.max(1, minimumRowCount);
  const targetColumnCount = Math.max(1, minimumColumnCount);
  const currentRowCount = getSheetRowCount(sheet);
  const currentColumnCount = getSheetColumnCount(sheet);

  if (currentColumnCount < targetColumnCount) {
    for (const row of sheet.cells) {
      row.push(...Array(targetColumnCount - currentColumnCount).fill(""));
    }
  }

  if (currentRowCount < targetRowCount) {
    sheet.cells.push(
      ...Array.from({ length: targetRowCount - currentRowCount }, () =>
        Array(Math.max(currentColumnCount, targetColumnCount)).fill(""),
      ),
    );
  }
}

function findSheetIndex(workbook: WorkbookState, sheetId: string): number {
  return workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
}

function getMutableSheet(
  workbook: WorkbookState,
  clonedSheetIds: Set<string>,
  requestedSheetId?: string,
): WorkbookSheet {
  const sheetId = requestedSheetId ?? workbook.activeSheetId;
  const sheetIndex = findSheetIndex(workbook, sheetId);

  if (sheetIndex < 0) {
    throw new Error(`Sheet "${sheetId}" was not found.`);
  }

  const currentSheet = workbook.sheets[sheetIndex];

  if (clonedSheetIds.has(sheetId)) {
    return currentSheet;
  }

  const clonedSheet: WorkbookSheet = {
    ...currentSheet,
    cells: currentSheet.cells.map((row) => [...row]),
    cellStyles: cloneCellStyles(currentSheet.cellStyles),
    columnWidths: cloneColumnWidths(currentSheet.columnWidths),
  };

  workbook.sheets[sheetIndex] = clonedSheet;
  clonedSheetIds.add(sheetId);

  return clonedSheet;
}

export function cloneWorkbookCellStyle(style: WorkbookCellStyle): WorkbookCellStyle {
  return {
    ...(style.backgroundColor ? { backgroundColor: style.backgroundColor } : {}),
    ...(style.bold ? { bold: true } : {}),
    ...(style.fontFamily ? { fontFamily: style.fontFamily } : {}),
    ...(style.fontSize !== undefined ? { fontSize: style.fontSize } : {}),
    ...(style.horizontalAlign ? { horizontalAlign: style.horizontalAlign } : {}),
    ...(style.italic ? { italic: true } : {}),
    ...(style.textColor ? { textColor: style.textColor } : {}),
    ...(style.wrapText ? { wrapText: true } : {}),
  };
}

function normalizeWorkbookCellStyle(style?: WorkbookCellStyle): WorkbookCellStyle | undefined {
  if (!style) {
    return undefined;
  }

  const normalized: WorkbookCellStyle = {};

  if (style.backgroundColor?.trim()) {
    normalized.backgroundColor = style.backgroundColor.trim();
  }

  if (style.bold === true) {
    normalized.bold = true;
  }

  if (style.fontFamily?.trim()) {
    normalized.fontFamily = style.fontFamily.trim();
  }

  if (style.fontSize !== undefined) {
    if (!Number.isFinite(style.fontSize) || style.fontSize < 6) {
      throw new Error("Cell style fontSize must be at least 6.");
    }

    normalized.fontSize = Math.floor(style.fontSize);
  }

  if (style.horizontalAlign !== undefined) {
    if (!["center", "left", "right"].includes(style.horizontalAlign)) {
      throw new Error("Cell style horizontalAlign is invalid.");
    }

    normalized.horizontalAlign = style.horizontalAlign;
  }

  if (style.italic === true) {
    normalized.italic = true;
  }

  if (style.textColor?.trim()) {
    normalized.textColor = style.textColor.trim();
  }

  if (style.wrapText === true) {
    normalized.wrapText = true;
  }

  return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function patchWorkbookCellStyle(
  baseStyle: WorkbookCellStyle | undefined,
  patch: WorkbookCellStylePatch | undefined,
): WorkbookCellStyle | undefined {
  const nextStyle = baseStyle ? cloneWorkbookCellStyle(baseStyle) : {};

  if (!patch) {
    return normalizeWorkbookCellStyle(nextStyle);
  }

  applyStringStylePatch(nextStyle, "backgroundColor", patch.backgroundColor);
  applyBooleanStylePatch(nextStyle, "bold", patch.bold);
  applyStringStylePatch(nextStyle, "fontFamily", patch.fontFamily);

  if (patch.fontSize !== undefined) {
    if (patch.fontSize === null) {
      delete nextStyle.fontSize;
    } else {
      nextStyle.fontSize = patch.fontSize;
    }
  }

  if (patch.horizontalAlign !== undefined) {
    if (patch.horizontalAlign === null) {
      delete nextStyle.horizontalAlign;
    } else {
      nextStyle.horizontalAlign = patch.horizontalAlign;
    }
  }

  applyBooleanStylePatch(nextStyle, "italic", patch.italic);
  applyStringStylePatch(nextStyle, "textColor", patch.textColor);
  applyBooleanStylePatch(nextStyle, "wrapText", patch.wrapText);

  return normalizeWorkbookCellStyle(nextStyle);
}

function applyStringStylePatch(
  style: WorkbookCellStyle,
  property: "backgroundColor" | "fontFamily" | "textColor",
  value: string | null | undefined,
) {
  if (value === undefined) {
    return;
  }

  if (value === null || value.trim().length === 0) {
    delete style[property];
    return;
  }

  style[property] = value.trim();
}

function applyBooleanStylePatch(
  style: WorkbookCellStyle,
  property: "bold" | "italic" | "wrapText",
  value: boolean | null | undefined,
) {
  if (value === undefined) {
    return;
  }

  if (value === true) {
    style[property] = true;
    return;
  }

  delete style[property];
}

function workbookCellStylesEqual(left?: WorkbookCellStyle, right?: WorkbookCellStyle): boolean {
  const normalizedLeft = normalizeWorkbookCellStyle(left);
  const normalizedRight = normalizeWorkbookCellStyle(right);

  return (
    normalizedLeft?.backgroundColor === normalizedRight?.backgroundColor &&
    normalizedLeft?.bold === normalizedRight?.bold &&
    normalizedLeft?.fontFamily === normalizedRight?.fontFamily &&
    normalizedLeft?.fontSize === normalizedRight?.fontSize &&
    normalizedLeft?.horizontalAlign === normalizedRight?.horizontalAlign &&
    normalizedLeft?.italic === normalizedRight?.italic &&
    normalizedLeft?.textColor === normalizedRight?.textColor &&
    normalizedLeft?.wrapText === normalizedRight?.wrapText
  );
}

function cloneCellStyles(
  styles: Record<string, WorkbookCellStyle>,
): Record<string, WorkbookCellStyle> {
  return Object.fromEntries(
    Object.entries(styles).map(([key, style]) => [key, cloneWorkbookCellStyle(style)]),
  );
}

function cloneColumnWidths(widths: Record<string, number>): Record<string, number> {
  return Object.fromEntries(
    Object.entries(widths)
      .map(([key, width]) => [key, normalizeColumnWidth(width)] as const)
      .filter(([, width]) => width !== DEFAULT_COLUMN_WIDTH),
  );
}

function normalizeColumnWidth(width: number): number {
  assertMinimumFiniteNumber(width, MIN_COLUMN_WIDTH, "Column width");
  assertMaximumFiniteNumber(width, MAX_COLUMN_WIDTH, "Column width");

  return Math.round(width);
}

function setColumnWidth(sheet: WorkbookSheet, columnIndex: number, width: number): boolean {
  const nextWidth = normalizeColumnWidth(width);
  const key = String(columnIndex);
  const currentWidth = sheet.columnWidths[key] ?? DEFAULT_COLUMN_WIDTH;

  if (currentWidth === nextWidth) {
    return false;
  }

  if (nextWidth === DEFAULT_COLUMN_WIDTH) {
    delete sheet.columnWidths[key];
    return true;
  }

  sheet.columnWidths[key] = nextWidth;
  return true;
}

function mapColumnWidths(
  widths: Record<string, number>,
  mapColumn: (columnIndex: number) => number | null,
): Record<string, number> {
  const nextWidths: Record<string, number> = {};

  for (const [key, width] of Object.entries(widths)) {
    const columnIndex = Number(key);

    if (!Number.isInteger(columnIndex) || columnIndex < 0) {
      continue;
    }

    const nextColumnIndex = mapColumn(columnIndex);

    if (nextColumnIndex === null) {
      continue;
    }

    const normalizedWidth = normalizeColumnWidth(width);

    if (normalizedWidth !== DEFAULT_COLUMN_WIDTH) {
      nextWidths[String(nextColumnIndex)] = normalizedWidth;
    }
  }

  return nextWidths;
}

function deleteColumnWidths(
  widths: Record<string, number>,
  deleteStart: number,
  count: number,
): Record<string, number> {
  const deleteEnd = deleteStart + count;

  return mapColumnWidths(widths, (columnIndex) => {
    if (columnIndex < deleteStart) {
      return columnIndex;
    }

    if (columnIndex >= deleteEnd) {
      return columnIndex - count;
    }

    return null;
  });
}

function insertColumnWidths(
  widths: Record<string, number>,
  insertAt: number,
  count: number,
): Record<string, number> {
  return mapColumnWidths(widths, (columnIndex) =>
    columnIndex >= insertAt ? columnIndex + count : columnIndex,
  );
}

function filterColumnWidthsInBounds(
  widths: Record<string, number>,
  columnCount: number,
): Record<string, number> {
  return mapColumnWidths(widths, (columnIndex) =>
    columnIndex >= columnCount ? null : columnIndex,
  );
}

function getCellKey(rowIndex: number, columnIndex: number): string {
  return `${rowIndex}:${columnIndex}`;
}

function parseCellKey(key: string): { columnIndex: number; rowIndex: number } {
  const [rowText, columnText] = key.split(":");

  return {
    columnIndex: Number.parseInt(columnText, 10),
    rowIndex: Number.parseInt(rowText, 10),
  };
}

function getSheetById(workbook: Pick<WorkbookState, "sheets">, sheetId: string): WorkbookSheet {
  const sheet = workbook.sheets.find((entry) => entry.id === sheetId);

  if (!sheet) {
    throw new Error(`Sheet "${sheetId}" was not found.`);
  }

  return sheet;
}

export function getSheetColumnCount(sheet: WorkbookSheet): number {
  return Math.max(1, sheet.cells[0]?.length ?? 0);
}

export function getSheetRowCount(sheet: WorkbookSheet): number {
  return Math.max(1, sheet.cells.length);
}

function matricesEqual(left: string[][], right: string[][]): boolean {
  if (left.length !== right.length) {
    return false;
  }

  for (let rowIndex = 0; rowIndex < left.length; rowIndex += 1) {
    const leftRow = left[rowIndex];
    const rightRow = right[rowIndex];

    if (leftRow.length !== rightRow.length) {
      return false;
    }

    for (let columnIndex = 0; columnIndex < leftRow.length; columnIndex += 1) {
      if (leftRow[columnIndex] !== rightRow[columnIndex]) {
        return false;
      }
    }
  }

  return true;
}

function setCellStyle(
  sheet: WorkbookSheet,
  rowIndex: number,
  columnIndex: number,
  style?: WorkbookCellStyle,
): boolean {
  const key = getCellKey(rowIndex, columnIndex);
  const currentStyle = sheet.cellStyles[key];

  if (workbookCellStylesEqual(currentStyle, style)) {
    return false;
  }

  if (!style) {
    delete sheet.cellStyles[key];
    return true;
  }

  sheet.cellStyles[key] = cloneWorkbookCellStyle(style);
  return true;
}

function clearCellStylesInRange(
  sheet: WorkbookSheet,
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number,
): boolean {
  const maxRow = getSheetRowCount(sheet);
  const maxColumn = getSheetColumnCount(sheet);
  const boundedStartRow = clampToRange(startRow, 0, maxRow);
  const boundedStartColumn = clampToRange(startColumn, 0, maxColumn);
  const endRow = Math.min(maxRow, boundedStartRow + Math.max(0, Math.floor(rowCount)));
  const endColumn = Math.min(maxColumn, boundedStartColumn + Math.max(0, Math.floor(columnCount)));
  let changed = false;

  for (let rowIndex = boundedStartRow; rowIndex < endRow; rowIndex += 1) {
    for (let columnIndex = boundedStartColumn; columnIndex < endColumn; columnIndex += 1) {
      const key = getCellKey(rowIndex, columnIndex);

      if (sheet.cellStyles[key]) {
        delete sheet.cellStyles[key];
        changed = true;
      }
    }
  }

  return changed;
}

function mapCellStyles(
  styles: Record<string, WorkbookCellStyle>,
  mapAddress: (
    rowIndex: number,
    columnIndex: number,
  ) => { columnIndex: number; rowIndex: number } | null,
): Record<string, WorkbookCellStyle> {
  const nextStyles: Record<string, WorkbookCellStyle> = {};

  for (const [key, style] of Object.entries(styles)) {
    const { columnIndex, rowIndex } = parseCellKey(key);
    const nextAddress = mapAddress(rowIndex, columnIndex);

    if (!nextAddress) {
      continue;
    }

    nextStyles[getCellKey(nextAddress.rowIndex, nextAddress.columnIndex)] =
      cloneWorkbookCellStyle(style);
  }

  return nextStyles;
}

function deleteColumnStyles(
  styles: Record<string, WorkbookCellStyle>,
  deleteStart: number,
  count: number,
): Record<string, WorkbookCellStyle> {
  const deleteEnd = deleteStart + count;

  return mapCellStyles(styles, (rowIndex, columnIndex) => {
    if (columnIndex < deleteStart) {
      return { columnIndex, rowIndex };
    }

    if (columnIndex >= deleteEnd) {
      return { columnIndex: columnIndex - count, rowIndex };
    }

    return null;
  });
}

function deleteRowStyles(
  styles: Record<string, WorkbookCellStyle>,
  deleteStart: number,
  count: number,
): Record<string, WorkbookCellStyle> {
  const deleteEnd = deleteStart + count;

  return mapCellStyles(styles, (rowIndex, columnIndex) => {
    if (rowIndex < deleteStart) {
      return { columnIndex, rowIndex };
    }

    if (rowIndex >= deleteEnd) {
      return { columnIndex, rowIndex: rowIndex - count };
    }

    return null;
  });
}

function insertColumnStyles(
  styles: Record<string, WorkbookCellStyle>,
  insertAt: number,
  count: number,
): Record<string, WorkbookCellStyle> {
  return mapCellStyles(styles, (rowIndex, columnIndex) => ({
    columnIndex: columnIndex >= insertAt ? columnIndex + count : columnIndex,
    rowIndex,
  }));
}

function insertRowStyles(
  styles: Record<string, WorkbookCellStyle>,
  insertAt: number,
  count: number,
): Record<string, WorkbookCellStyle> {
  return mapCellStyles(styles, (rowIndex, columnIndex) => ({
    columnIndex,
    rowIndex: rowIndex >= insertAt ? rowIndex + count : rowIndex,
  }));
}

function filterCellStylesInBounds(
  styles: Record<string, WorkbookCellStyle>,
  rowCount: number,
  columnCount: number,
): Record<string, WorkbookCellStyle> {
  return mapCellStyles(styles, (rowIndex, columnIndex) => {
    if (rowIndex < 0 || columnIndex < 0 || rowIndex >= rowCount || columnIndex >= columnCount) {
      return null;
    }

    return { columnIndex, rowIndex };
  });
}

function resizeMatrix(matrix: string[][], rowCount: number, columnCount: number): string[][] {
  const nextRows = Math.max(1, Math.floor(rowCount));
  const nextColumns = Math.max(1, Math.floor(columnCount));

  return Array.from({ length: nextRows }, (_, rowIndex) => {
    const sourceRow = matrix[rowIndex] ?? [];

    return Array.from({ length: nextColumns }, (_, columnIndex) => sourceRow[columnIndex] ?? "");
  });
}

function registerSheetId(sheetId: string) {
  const match = /^sheet-(\d+)$/.exec(sheetId);

  if (!match) {
    return;
  }

  const sequence = Number.parseInt(match[1], 10) + 1;

  if (!Number.isNaN(sequence)) {
    nextSheetIdSequence = Math.max(nextSheetIdSequence, sequence);
  }
}
