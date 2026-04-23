import {
  getColumnTitle,
  DEFAULT_CHART_LAYOUT_HEIGHT,
  DEFAULT_CHART_LAYOUT_WIDTH,
  getWorkbookChartValidationIssues,
  parseCellReference,
  type WorkbookChart,
  type WorkbookChartRange,
  type WorkbookChartSpec,
  type WorkbookChartType,
  type WorkbookChartValidationIssue,
  type WorkbookSummary,
  type WorkbookTransactionOperation,
  type UsedRangeResult,
} from "./workbook-core";

export type ChartEditorWindowRequest =
  | {
      mode: "create";
      sheetId?: string;
    }
  | {
      chartId: string;
      mode: "edit";
    };

export interface ChartEditorFormState {
  categoryDimension: string;
  chartType: WorkbookChartType;
  name: string;
  nameDimension: string;
  seriesLayoutBy: "column" | "row";
  smooth: boolean;
  sourceRange: string;
  sourceHeader: boolean;
  stacked: boolean;
  valueDimension: string;
  valueDimensions: string;
}

export function createChartEditorFormState(
  request: ChartEditorWindowRequest,
  summary: WorkbookSummary,
  usedRange: UsedRangeResult,
  chart?: WorkbookChart,
): ChartEditorFormState {
  if (request.mode === "edit" && chart) {
    return createChartEditorFormStateFromChart(chart);
  }

  const sheet =
    summary.sheets.find(
      (entry) =>
        entry.id === (request.mode === "create" ? (request.sheetId ?? "") : ""),
    ) ??
    summary.sheets.find((entry) => entry.id === summary.activeSheetId) ??
    summary.sheets[0];
  const minimumRowCount = Math.min(sheet?.rowCount ?? 2, 6);
  const rowCount =
    usedRange.rowCount > 0
      ? Math.max(2, usedRange.rowCount)
      : Math.max(2, minimumRowCount);
  const columnCount =
    usedRange.columnCount > 0 ? Math.max(2, usedRange.columnCount) : 2;

  return {
    categoryDimension: "0",
    chartType: "scatter",
    name: "New Scatter Chart",
    nameDimension: "0",
    seriesLayoutBy: "column",
    smooth: false,
    sourceRange: formatChartEditorRange({
      columnCount,
      rowCount,
      startColumn: usedRange.startColumn,
      startRow: usedRange.startRow,
    }),
    sourceHeader: true,
    stacked: false,
    valueDimension: "1",
    valueDimensions: "1",
  };
}

export function createChartEditorFormStateFromChart(
  chart: WorkbookChart,
): ChartEditorFormState {
  const range = chart.spec.source.range;

  return {
    categoryDimension:
      chart.spec.family === "cartesian"
        ? `${chart.spec.categoryDimension}`
        : "0",
    chartType: chart.spec.chartType,
    name: chart.name,
    nameDimension:
      chart.spec.family === "pie" ? `${chart.spec.nameDimension}` : "0",
    seriesLayoutBy: chart.spec.source.seriesLayoutBy,
    smooth:
      chart.spec.family === "cartesian" ? (chart.spec.smooth ?? false) : false,
    sourceRange: formatChartEditorRange(range),
    sourceHeader: chart.spec.source.sourceHeader,
    stacked:
      chart.spec.family === "cartesian" ? (chart.spec.stacked ?? false) : false,
    valueDimension:
      chart.spec.family === "pie" ? `${chart.spec.valueDimension}` : "1",
    valueDimensions:
      chart.spec.family === "cartesian"
        ? chart.spec.valueDimensions.join(", ")
        : "1",
  };
}

export function buildChartEditorSpec(
  sheetId: string,
  state: ChartEditorFormState,
): WorkbookChartSpec {
  const range = parseChartEditorRange(state.sourceRange, sheetId);
  const source = {
    range,
    seriesLayoutBy: state.seriesLayoutBy,
    sourceHeader: state.sourceHeader,
  } as const;

  if (state.chartType === "pie") {
    return {
      chartType: "pie",
      family: "pie",
      nameDimension: parseIntegerField(state.nameDimension),
      source,
      valueDimension: parseIntegerField(state.valueDimension),
    };
  }

  return {
    categoryDimension: parseIntegerField(state.categoryDimension),
    chartType: state.chartType,
    family: "cartesian",
    ...(state.chartType === "line" || state.chartType === "area"
      ? {
          smooth: state.smooth,
        }
      : {}),
    source,
    ...(state.chartType === "bar" || state.chartType === "area"
      ? {
          stacked: state.stacked,
        }
      : {}),
    valueDimensions: parseIntegerListField(state.valueDimensions),
  };
}

export function buildChartEditorOperations(
  request: ChartEditorWindowRequest,
  sheetId: string,
  state: ChartEditorFormState,
): WorkbookTransactionOperation[] {
  const spec = buildChartEditorSpec(sheetId, state);
  const name = state.name.trim();

  if (request.mode === "edit") {
    return [
      {
        chartId: request.chartId,
        name,
        type: "renameChart",
      },
      {
        chartId: request.chartId,
        spec,
        type: "setChartSpec",
      },
    ];
  }

  return [
    {
      name,
      spec,
      type: "addChart",
    },
  ];
}

export function getChartEditorValidationIssues(
  state: ChartEditorFormState,
  sheetId: string,
  summary: WorkbookSummary,
  chartId = "draft-chart",
): WorkbookChartValidationIssue[] {
  const name = state.name.trim();

  if (name.length === 0) {
    return [
      {
        code: "INVALID_DIMENSION",
        message: "Chart name is required.",
      },
    ];
  }

  try {
    return getWorkbookChartValidationIssues(
      {
        id: chartId,
        layout: {
          height: DEFAULT_CHART_LAYOUT_HEIGHT,
          offsetX: 0,
          offsetY: 0,
          startColumn: 0,
          startRow: 0,
          width: DEFAULT_CHART_LAYOUT_WIDTH,
          zIndex: 0,
        },
        name,
        sheetId,
        spec: buildChartEditorSpec(sheetId, state),
      },
      summary.sheets.map((entry) => ({
        columnCount: entry.columnCount,
        id: entry.id,
        rowCount: entry.rowCount,
      })),
    );
  } catch (error) {
    return [
      {
        code: "INVALID_DIMENSION",
        message:
          error instanceof Error
            ? error.message
            : "Chart settings are invalid.",
      },
    ];
  }
}

export function getChartEditorSheetId(
  request: ChartEditorWindowRequest,
  summary: WorkbookSummary,
  chart?: WorkbookChart,
): string {
  if (request.mode === "edit" && chart) {
    return chart.sheetId;
  }

  return request.mode === "create" && request.sheetId
    ? request.sheetId
    : summary.activeSheetId;
}

function parseIntegerField(value: string): number {
  const parsed = Number.parseInt(value.trim(), 10);

  if (Number.isNaN(parsed)) {
    throw new Error(`"${value}" is not a valid integer.`);
  }

  return parsed;
}

function parseIntegerListField(value: string): number[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item.length > 0)
    .map((item) => parseIntegerField(item));
}

export function formatChartEditorRange(
  range: Pick<
    WorkbookChartRange,
    "columnCount" | "rowCount" | "startColumn" | "startRow"
  >,
): string {
  if (range.columnCount < 1 || range.rowCount < 1) {
    return "";
  }

  const endColumn = range.startColumn + range.columnCount - 1;
  const endRow = range.startRow + range.rowCount;

  return `${getColumnTitle(range.startColumn)}${range.startRow + 1}:${getColumnTitle(
    endColumn,
  )}${endRow}`;
}

function parseChartEditorRange(
  value: string,
  sheetId: string,
): WorkbookChartRange {
  const normalizedValue = value.trim();
  const parts = normalizedValue.split(":").map((part) => part.trim());

  if (
    parts.length < 1 ||
    parts.length > 2 ||
    parts.some((part) => part === "")
  ) {
    throw new Error(`"${value}" is not a valid cell range.`);
  }

  const start = parseCellReference(parts[0]);
  const end = parseCellReference(parts[1] ?? parts[0]);
  const startRow = Math.min(start.rowIndex, end.rowIndex);
  const startColumn = Math.min(start.columnIndex, end.columnIndex);
  const endRow = Math.max(start.rowIndex, end.rowIndex);
  const endColumn = Math.max(start.columnIndex, end.columnIndex);

  return {
    columnCount: endColumn - startColumn + 1,
    rowCount: endRow - startRow + 1,
    sheetId,
    startColumn,
    startRow,
  };
}
