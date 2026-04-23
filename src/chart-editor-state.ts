import {
  getWorkbookChartValidationIssues,
  type WorkbookChart,
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
  columnCount: string;
  name: string;
  nameDimension: string;
  rowCount: string;
  seriesLayoutBy: "column" | "row";
  smooth: boolean;
  sourceHeader: boolean;
  stacked: boolean;
  startColumn: string;
  startRow: string;
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
      (entry) => entry.id === (request.mode === "create" ? request.sheetId ?? "" : ""),
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
    columnCount: `${columnCount}`,
    name: "New Scatter Chart",
    nameDimension: "0",
    rowCount: `${rowCount}`,
    seriesLayoutBy: "column",
    smooth: false,
    sourceHeader: true,
    stacked: false,
    startColumn: `${usedRange.startColumn}`,
    startRow: `${usedRange.startRow}`,
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
    columnCount: `${range.columnCount}`,
    name: chart.name,
    nameDimension:
      chart.spec.family === "pie" ? `${chart.spec.nameDimension}` : "0",
    rowCount: `${range.rowCount}`,
    seriesLayoutBy: chart.spec.source.seriesLayoutBy,
    smooth: chart.spec.family === "cartesian" ? chart.spec.smooth ?? false : false,
    sourceHeader: chart.spec.source.sourceHeader,
    stacked:
      chart.spec.family === "cartesian" ? chart.spec.stacked ?? false : false,
    startColumn: `${range.startColumn}`,
    startRow: `${range.startRow}`,
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
  const source = {
    range: {
      columnCount: parseIntegerField(state.columnCount),
      rowCount: parseIntegerField(state.rowCount),
      sheetId,
      startColumn: parseIntegerField(state.startColumn),
      startRow: parseIntegerField(state.startRow),
    },
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
          error instanceof Error ? error.message : "Chart settings are invalid.",
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
