export const DEFAULT_INITIAL_ROWS = 200;
export const DEFAULT_INITIAL_COLUMNS = 50;
export const DEFAULT_SHEET_NAME = "Sheet 1";

export interface ControlServerInfo {
  host: string;
  port: number;
  protocol: "jsonl";
}

export interface WorkbookSheet {
  id: string;
  name: string;
  cells: string[][];
  sourceFilePath?: string;
}

export interface WorkbookState {
  charts: WorkbookChart[];
  documentFilePath?: string;
  hasUnsavedChanges: boolean;
  version: number;
  activeSheetId: string;
  nextChartNumber: number;
  nextSheetNumber: number;
  sheets: WorkbookSheet[];
}

export interface SheetSummary {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
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
}

export const WORKBOOK_CHART_TYPES = [
  "bar",
  "line",
  "area",
  "pie",
  "scatter",
] as const;

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

export type WorkbookChartSpec =
  | WorkbookChartCartesianSpec
  | WorkbookChartPieSpec;

export interface WorkbookChart {
  id: string;
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
      type: "clearRange";
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
      type: "setActiveSheet";
      sheetId: string;
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
      type: "setRange";
      sheetId?: string;
      startColumn: number;
      startRow: number;
      values: string[][];
    };

export interface ApplyTransactionRequest {
  dryRun?: boolean;
  operations: WorkbookTransactionOperation[];
}

export interface ApplyTransactionResult {
  changed: boolean;
  summary: WorkbookSummary;
  version: number;
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

  return Array.from({ length: normalizedRowCount }, () =>
    Array(normalizedColumnCount).fill(""),
  );
}

export function normalizeSheet(rows: string[][]): string[][] {
  const rowCount = Math.max(rows.length, 1);
  const columnCount = Math.max(1, ...rows.map((row) => row.length));

  return Array.from({ length: rowCount }, (_, rowIndex) => {
    const sourceRow = rows[rowIndex] ?? [];

    return Array.from(
      { length: columnCount },
      (_, columnIndex) => sourceRow[columnIndex] ?? "",
    );
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
  if (
    value.includes(delimiter) ||
    value.includes('"') ||
    /[\r\n]/.test(value)
  ) {
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
    .map((row) =>
      row.map((value) => escapeDelimitedValue(value ?? "", "\t")).join("\t"),
    )
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
    nextSheetNumber: 2,
    sheets: [defaultSheet],
    version: 0,
  };
}

export function getWorkbookSummary(workbook: WorkbookState): WorkbookSummary {
  const activeSheet = getSheetById(workbook, workbook.activeSheetId);
  const chartSheets: WorkbookChartSheetReference[] = workbook.sheets.map(
    (sheet) => ({
      columnCount: getSheetColumnCount(sheet),
      id: sheet.id,
      rowCount: getSheetRowCount(sheet),
    }),
  );

  return {
    activeSheetId: activeSheet.id,
    activeSheetName: activeSheet.name,
    charts: workbook.charts.map((chart) =>
      createWorkbookChartSummary(chart, chartSheets),
    ),
    documentFilePath: workbook.documentFilePath,
    hasUnsavedChanges: workbook.hasUnsavedChanges,
    sheets: workbook.sheets.map((sheet) => ({
      columnCount: getSheetColumnCount(sheet),
      id: sheet.id,
      name: sheet.name,
      rowCount: getSheetRowCount(sheet),
      sourceFilePath: sheet.sourceFilePath,
    })),
    version: workbook.version,
  };
}

export function cloneWorkbookChart(chart: WorkbookChart): WorkbookChart {
  return {
    ...chart,
    spec:
      chart.spec.family === "cartesian"
        ? {
            ...chart.spec,
            source: {
              ...chart.spec.source,
              range: {
                ...chart.spec.source.range,
              },
            },
            valueDimensions: [...chart.spec.valueDimensions],
          }
        : {
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

export function getWorkbookChartById(
  workbook: WorkbookState,
  chartId: string,
): WorkbookChart {
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

export function getSheetRange(
  workbook: WorkbookState,
  request: SheetRangeRequest,
): SheetRangeResult {
  const sheet = getSheetById(
    workbook,
    request.sheetId ?? workbook.activeSheetId,
  );
  const rowCount = getSheetRowCount(sheet);
  const columnCount = getSheetColumnCount(sheet);
  const startRow = clampToRange(request.startRow, 0, rowCount);
  const startColumn = clampToRange(request.startColumn, 0, columnCount);
  const requestedRowCount = Math.max(0, Math.floor(request.rowCount));
  const requestedColumnCount = Math.max(0, Math.floor(request.columnCount));
  const boundedRowCount = Math.max(
    0,
    Math.min(requestedRowCount, rowCount - startRow),
  );
  const boundedColumnCount = Math.max(
    0,
    Math.min(requestedColumnCount, columnCount - startColumn),
  );

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

export function getSheetCsv(workbook: WorkbookState, sheetId?: string): string {
  return serializeCsv(
    getSheetById(workbook, sheetId ?? workbook.activeSheetId),
  );
}

export function getSheetUsedRange(
  workbook: WorkbookState,
  sheetId?: string,
): UsedRangeResult {
  return getUsedRange(
    getSheetById(workbook, sheetId ?? workbook.activeSheetId),
  );
}

export function getWorkbookSheet(
  workbook: WorkbookState,
  sheetId?: string,
): WorkbookSheet {
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

export function getWorkbookChartDimensionCount(
  chart: Pick<WorkbookChart, "spec">,
): number {
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

  if (range.sheetId !== chart.sheetId) {
    issues.push({
      code: "CROSS_SHEET_SOURCE",
      message:
        "Chart sources must stay on the same sheet as the chart in v1 chart support.",
    });
  }

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
      message:
        "Chart source ranges require non-negative integer coordinates and sizes.",
    });
  }

  if (range.rowCount < 1 || range.columnCount < 1) {
    issues.push({
      code: "EMPTY_RANGE",
      message:
        "Chart source ranges must contain at least one row and one column.",
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
        message:
          "Cartesian charts require at least one value dimension in the shared contract.",
      });
    }

    if (
      new Set(chart.spec.valueDimensions).size !==
      chart.spec.valueDimensions.length
    ) {
      issues.push({
        code: "REPEATED_VALUE_DIMENSION",
        message: "Cartesian chart value dimensions must not repeat.",
      });
    }

    const dimensionsToCheck = [
      chart.spec.categoryDimension,
      ...chart.spec.valueDimensions,
    ];

    if (
      dimensionsToCheck.some(
        (dimension) =>
          !Number.isInteger(dimension) ||
          dimension < 0 ||
          dimension >= dimensionCount,
      )
    ) {
      issues.push({
        code: "INVALID_DIMENSION",
        message:
          "Cartesian chart dimensions must resolve inside the source table.",
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
  return getWorkbookChartValidationIssues(chart, sheets).length === 0
    ? "ok"
    : "invalid";
}

export function createWorkbookChartSummary(
  chart: WorkbookChart,
  sheets: readonly WorkbookChartSheetReference[],
): WorkbookChartSummary {
  return {
    chartType: chart.spec.chartType,
    id: chart.id,
    name: chart.name,
    sheetId: chart.sheetId,
    status: getWorkbookChartStatus(chart, sheets),
  };
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
  const nextRangeAxis = adjustRangeAxisForDeletion(
    range.startRow,
    range.rowCount,
    rowIndex,
    count,
  );

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
  };
  const clonedSheetIds = new Set<string>();
  let changed = false;

  for (const operation of request.operations) {
    switch (operation.type) {
      case "addSheet": {
        const sheetName =
          operation.name?.trim() || `Sheet ${nextState.nextSheetNumber}`;
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
        nextState.nextSheetNumber += 1;

        if (operation.activate ?? true) {
          nextState.activeSheetId = sheetId;
        }

        changed = true;
        break;
      }

      case "clearRange": {
        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const maxRow = getSheetRowCount(sheet);
        const maxColumn = getSheetColumnCount(sheet);
        const startRow = clampToRange(operation.startRow, 0, maxRow);
        const startColumn = clampToRange(operation.startColumn, 0, maxColumn);
        const endRow = Math.min(
          maxRow,
          startRow + Math.max(0, Math.floor(operation.rowCount)),
        );
        const endColumn = Math.min(
          maxColumn,
          startColumn + Math.max(0, Math.floor(operation.columnCount)),
        );

        for (let rowIndex = startRow; rowIndex < endRow; rowIndex += 1) {
          const row = sheet.cells[rowIndex];

          for (
            let columnIndex = startColumn;
            columnIndex < endColumn;
            columnIndex += 1
          ) {
            if (row[columnIndex] === "") {
              continue;
            }

            row[columnIndex] = "";
            changed = true;
          }
        }

        break;
      }

      case "deleteColumns": {
        assertPositiveCount(operation.count, "Column delete count");

        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const currentColumnCount = getSheetColumnCount(sheet);
        const deleteStart = clampToRange(
          operation.columnIndex,
          0,
          currentColumnCount,
        );
        const requestedDeleteCount = Math.min(
          operation.count,
          currentColumnCount - deleteStart,
        );

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

        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartForDeletedColumns(
            chart,
            sheet.id,
            deleteStart,
            requestedDeleteCount,
          ),
        );
        changed = true;
        break;
      }

      case "deleteRows": {
        assertPositiveCount(operation.count, "Row delete count");

        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const currentRowCount = getSheetRowCount(sheet);
        const deleteStart = clampToRange(
          operation.rowIndex,
          0,
          currentRowCount,
        );
        const requestedDeleteCount = Math.min(
          operation.count,
          currentRowCount - deleteStart,
        );

        if (requestedDeleteCount === 0) {
          break;
        }

        if (requestedDeleteCount >= currentRowCount) {
          sheet.cells.splice(
            0,
            sheet.cells.length,
            Array(getSheetColumnCount(sheet)).fill(""),
          );
        } else {
          sheet.cells.splice(deleteStart, requestedDeleteCount);
        }

        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartForDeletedRows(
            chart,
            sheet.id,
            deleteStart,
            requestedDeleteCount,
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

        if (nextState.activeSheetId === deletedSheet.id) {
          const nextActiveSheet =
            (operation.nextActiveSheetId
              ? nextState.sheets.find(
                  (sheet) => sheet.id === operation.nextActiveSheetId,
                )
              : undefined) ?? nextState.sheets[Math.max(0, deleteIndex - 1)];

          nextState.activeSheetId = nextActiveSheet.id;
        }

        changed = true;
        break;
      }

      case "insertColumns": {
        assertPositiveCount(operation.count, "Column insert count");

        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const insertAt = clampToRange(
          operation.columnIndex,
          0,
          getSheetColumnCount(sheet),
        );

        for (const row of sheet.cells) {
          row.splice(insertAt, 0, ...Array(operation.count).fill(""));
        }

        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartForInsertedColumns(
            chart,
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

        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const insertAt = clampToRange(
          operation.rowIndex,
          0,
          getSheetRowCount(sheet),
        );
        const columnCount = getSheetColumnCount(sheet);

        sheet.cells.splice(
          insertAt,
          0,
          ...Array.from({ length: operation.count }, () =>
            Array(columnCount).fill(""),
          ),
        );
        nextState.charts = nextState.charts.map((chart) =>
          adjustWorkbookChartForInsertedRows(
            chart,
            sheet.id,
            insertAt,
            operation.count,
          ),
        );
        changed = true;
        break;
      }

      case "renameSheet": {
        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const nextName = operation.name.trim();

        if (nextName.length === 0 || sheet.name === nextName) {
          break;
        }

        sheet.name = nextName;
        changed = true;
        break;
      }

      case "replaceSheet": {
        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const nextCells = normalizeSheet(operation.rows);

        if (!matricesEqual(sheet.cells, nextCells)) {
          sheet.cells = nextCells;
          changed = true;
        }

        if (operation.name?.trim() && operation.name.trim() !== sheet.name) {
          sheet.name = operation.name.trim();
          changed = true;
        }

        if (sheet.sourceFilePath !== operation.sourceFilePath) {
          sheet.sourceFilePath = operation.sourceFilePath;
          changed = true;
        }

        break;
      }

      case "replaceSheetFromCsv": {
        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const nextCells = parseCsv(operation.content);

        if (!matricesEqual(sheet.cells, nextCells)) {
          sheet.cells = nextCells;
          changed = true;
        }

        if (operation.name?.trim() && operation.name.trim() !== sheet.name) {
          sheet.name = operation.name.trim();
          changed = true;
        }

        if (sheet.sourceFilePath !== operation.sourceFilePath) {
          sheet.sourceFilePath = operation.sourceFilePath;
          changed = true;
        }
        break;
      }

      case "resizeSheet": {
        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const targetRowCount = Math.max(1, Math.floor(operation.rowCount));
        const targetColumnCount = Math.max(
          1,
          Math.floor(operation.columnCount),
        );

        if (
          targetRowCount === getSheetRowCount(sheet) &&
          targetColumnCount === getSheetColumnCount(sheet)
        ) {
          break;
        }

        sheet.cells = resizeMatrix(
          sheet.cells,
          targetRowCount,
          targetColumnCount,
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

      case "setSheetSourceFile": {
        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );

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

        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        ensureSheetSize(
          sheet,
          operation.rowIndex + 1,
          operation.columnIndex + 1,
        );

        if (
          sheet.cells[operation.rowIndex][operation.columnIndex] ===
          operation.value
        ) {
          break;
        }

        sheet.cells[operation.rowIndex][operation.columnIndex] =
          operation.value;
        changed = true;
        break;
      }

      case "setRange": {
        if (operation.values.length === 0) {
          break;
        }

        assertNonNegativeIndex(operation.startRow, "Start row");
        assertNonNegativeIndex(operation.startColumn, "Start column");

        const sheet = getMutableSheet(
          nextState,
          clonedSheetIds,
          operation.sheetId,
        );
        const maxColumnCount = Math.max(
          0,
          ...operation.values.map((row) => row.length),
        );

        if (maxColumnCount === 0) {
          break;
        }

        ensureSheetSize(
          sheet,
          operation.startRow + operation.values.length,
          operation.startColumn + maxColumnCount,
        );

        for (
          let rowOffset = 0;
          rowOffset < operation.values.length;
          rowOffset += 1
        ) {
          const sourceRow = operation.values[rowOffset];
          const targetRow = sheet.cells[operation.startRow + rowOffset];

          for (
            let columnOffset = 0;
            columnOffset < sourceRow.length;
            columnOffset += 1
          ) {
            const nextValue = sourceRow[columnOffset] ?? "";
            const targetColumn = operation.startColumn + columnOffset;

            if (targetRow[targetColumn] === nextValue) {
              continue;
            }

            targetRow[targetColumn] = nextValue;
            changed = true;
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

function clampToRange(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(Math.floor(value), max));
}

function updateWorkbookChartRange(
  chart: WorkbookChart,
  range: WorkbookChartRange,
): WorkbookChart {
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

  const survivorCountBeforeDelete = Math.max(
    0,
    Math.min(end, deleteStart) - start,
  );
  const survivorCountAfterDelete = Math.max(
    0,
    end - Math.max(start, deleteEnd),
  );
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

function createWorkbookSheet(
  name: string,
  rowCount: number,
  columnCount: number,
  id = createSheetId(),
): WorkbookSheet {
  registerSheetId(id);

  return {
    cells: createSheet(rowCount, columnCount),
    id,
    name,
  };
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
  };

  workbook.sheets[sheetIndex] = clonedSheet;
  clonedSheetIds.add(sheetId);

  return clonedSheet;
}

function getSheetById(workbook: WorkbookState, sheetId: string): WorkbookSheet {
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

function resizeMatrix(
  matrix: string[][],
  rowCount: number,
  columnCount: number,
): string[][] {
  const nextRows = Math.max(1, Math.floor(rowCount));
  const nextColumns = Math.max(1, Math.floor(columnCount));

  return Array.from({ length: nextRows }, (_, rowIndex) => {
    const sourceRow = matrix[rowIndex] ?? [];

    return Array.from(
      { length: nextColumns },
      (_, columnIndex) => sourceRow[columnIndex] ?? "",
    );
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
