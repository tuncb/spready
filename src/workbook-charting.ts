import {
  getCellEvaluation,
  type CellEvaluation,
  type SheetEvaluationSnapshot,
} from "./formula-engine";
import {
  getWorkbookChartStatus,
  getWorkbookChartValidationIssues,
  type WorkbookChart,
  type WorkbookChartCartesianSpec,
  type WorkbookChartDimensionType,
  type WorkbookChartPieSpec,
  type WorkbookChartPreview,
  type WorkbookChartPreviewDataset,
  type WorkbookChartPreviewDimension,
  type WorkbookChartSheetReference,
  type WorkbookSheet,
} from "./workbook-core";

type WorkbookCartesianChart = WorkbookChart & {
  spec: WorkbookChartCartesianSpec;
};

type WorkbookPieChart = WorkbookChart & {
  spec: WorkbookChartPieSpec;
};

const CARTESIAN_GRID_BOTTOM = 18;
const CARTESIAN_GRID_LEFT_MIN = 56;
const CARTESIAN_GRID_RIGHT = 16;
const CARTESIAN_GRID_TOP = 16;
const Y_AXIS_NAME_GAP_MIN = 42;
const Y_AXIS_TICK_LABEL_MARGIN = 8;
const Y_AXIS_NAME_PADDING = 18;
const Y_AXIS_NAME_GRID_PADDING = 12;
const AXIS_LABEL_FONT_SIZE = 12;

export function buildWorkbookChartPreview(
  chart: WorkbookChart,
  sheet: WorkbookSheet | undefined,
  snapshot: SheetEvaluationSnapshot | undefined,
  sheets: readonly WorkbookChartSheetReference[],
): WorkbookChartPreview {
  const status = getWorkbookChartStatus(chart, sheets);
  const validationIssues = getWorkbookChartValidationIssues(chart, sheets);
  const warningSet = new Set(validationIssues.map((issue) => issue.message));

  if (!sheet || !snapshot || status === "invalid") {
    return {
      chart,
      dataset: createEmptyPreviewDataset(),
      option: createInvalidChartOption(chart),
      status,
      validationIssues,
      warnings: [...warningSet],
    };
  }

  const normalizedTable = normalizeChartSourceMatrix(chart, sheet, snapshot);
  const dataset = buildPreviewDataset(chart, normalizedTable, warningSet);

  return {
    chart,
    dataset,
    option: buildChartOption(chart, dataset),
    status,
    validationIssues,
    warnings: [...warningSet],
  };
}

function createEmptyPreviewDataset(): WorkbookChartPreviewDataset {
  return {
    dimensions: [],
    seriesLayoutBy: "column",
    source: [],
    sourceHeader: false,
  };
}

function createInvalidChartOption(chart: WorkbookChart): Record<string, unknown> {
  return {
    series: [],
    title: {
      text: chart.name,
    },
  };
}

function normalizeChartSourceMatrix(
  chart: WorkbookChart,
  sheet: WorkbookSheet,
  snapshot: SheetEvaluationSnapshot,
): CellEvaluation[][] {
  const { range, seriesLayoutBy } = chart.spec.source;
  const maxRowCount = Math.max(0, sheet.cells.length - range.startRow);
  const maxColumnCount = Math.max(0, (sheet.cells[0]?.length ?? 0) - range.startColumn);
  const rowCount = Math.max(0, Math.min(range.rowCount, maxRowCount));
  const columnCount = Math.max(0, Math.min(range.columnCount, maxColumnCount));
  const matrix = Array.from({ length: rowCount }, (_, rowOffset) =>
    Array.from({ length: columnCount }, (_, columnOffset) =>
      getCellEvaluation(snapshot, range.startRow + rowOffset, range.startColumn + columnOffset),
    ),
  );

  return seriesLayoutBy === "row" ? transposeMatrix(matrix) : matrix;
}

function transposeMatrix(matrix: CellEvaluation[][]): CellEvaluation[][] {
  const rowCount = matrix.length;
  const columnCount = Math.max(0, ...matrix.map((row) => row.length));

  return Array.from({ length: columnCount }, (_, columnIndex) =>
    Array.from({ length: rowCount }, (_, rowIndex) => matrix[rowIndex]?.[columnIndex]),
  );
}

function buildPreviewDataset(
  chart: WorkbookChart,
  matrix: CellEvaluation[][],
  warningSet: Set<string>,
): WorkbookChartPreviewDataset {
  if (matrix.length === 0 || matrix[0]?.length === 0) {
    return createEmptyPreviewDataset();
  }

  const sourceHeader = chart.spec.source.sourceHeader;
  const headerRow = sourceHeader ? matrix[0] : undefined;
  const dataRows = sourceHeader ? matrix.slice(1) : matrix;
  const columnCount = matrix[0]?.length ?? 0;
  const dimensions = Array.from({ length: columnCount }, (_, columnIndex) =>
    createPreviewDimension(headerRow?.[columnIndex], dataRows, columnIndex),
  );
  const source = dataRows.map((row) =>
    Array.from({ length: columnCount }, (_, columnIndex) =>
      toPreviewValue(row[columnIndex], warningSet),
    ),
  );

  return {
    dimensions,
    seriesLayoutBy: "column",
    source: sourceHeader ? [dimensions.map((dimension) => dimension.name), ...source] : source,
    sourceHeader,
  };
}

function createPreviewDimension(
  headerCell: CellEvaluation | undefined,
  dataRows: CellEvaluation[][],
  columnIndex: number,
): WorkbookChartPreviewDimension {
  const name = getDimensionName(headerCell, columnIndex);
  const previewValues = dataRows.map((row) => toPreviewValue(row[columnIndex]));

  return {
    name,
    type: inferDimensionType(previewValues),
  };
}

function getDimensionName(headerCell: CellEvaluation | undefined, columnIndex: number): string {
  if (!headerCell) {
    return `Dimension ${columnIndex + 1}`;
  }

  const headerValue = headerCell.display.trim();

  return headerValue.length > 0 ? headerValue : `Dimension ${columnIndex + 1}`;
}

function toPreviewValue(
  cell: CellEvaluation | undefined,
  warningSet?: Set<string>,
): string | number | null {
  if (!cell) {
    return null;
  }

  switch (cell.value.type) {
    case "blank":
      return null;
    case "boolean":
      return cell.display;
    case "error":
      warningSet?.add(
        "Chart preview skipped one or more formula errors by converting them to null values.",
      );
      return null;
    case "number":
      return cell.value.value;
    case "text":
      return cell.value.value;
  }
}

function inferDimensionType(values: Array<string | number | null>): WorkbookChartDimensionType {
  const nonNullValues = values.filter((value) => value !== null);

  if (nonNullValues.length === 0) {
    return "ordinal";
  }

  if (nonNullValues.every((value) => typeof value === "number")) {
    return "number";
  }

  if (nonNullValues.every((value) => typeof value === "string" && looksLikeTimeString(value))) {
    return "time";
  }

  return "ordinal";
}

function looksLikeTimeString(value: string): boolean {
  if (!/[-/:T]/.test(value)) {
    return false;
  }

  return !Number.isNaN(Date.parse(value));
}

function buildChartOption(
  chart: WorkbookChart,
  dataset: WorkbookChartPreviewDataset,
): Record<string, unknown> {
  if (isCartesianChart(chart)) {
    return buildCartesianChartOption(chart, dataset);
  }

  return buildPieChartOption(chart as WorkbookPieChart, dataset);
}

function buildCartesianChartOption(
  chart: WorkbookCartesianChart,
  dataset: WorkbookChartPreviewDataset,
): Record<string, unknown> {
  const categoryKey = getDimensionEncodeKey(dataset, chart.spec.categoryDimension);
  const categoryDimension = dataset.dimensions[chart.spec.categoryDimension]?.type ?? "ordinal";
  const seriesType = chart.spec.chartType === "area" ? "line" : chart.spec.chartType;
  const xAxisName = getDimensionLabel(dataset, chart.spec.categoryDimension);
  const yAxisName = getValueAxisLabel(dataset, chart.spec.valueDimensions);
  const yAxisNameGap = getValueAxisNameGap(chart, dataset);

  return {
    dataset: {
      dimensions: dataset.dimensions,
      source: dataset.source,
      sourceHeader: dataset.sourceHeader,
    },
    grid: {
      bottom: CARTESIAN_GRID_BOTTOM,
      containLabel: true,
      left: Math.max(CARTESIAN_GRID_LEFT_MIN, yAxisNameGap + Y_AXIS_NAME_GRID_PADDING),
      right: CARTESIAN_GRID_RIGHT,
      top: CARTESIAN_GRID_TOP,
    },
    legend: {
      show: chart.spec.valueDimensions.length > 1,
    },
    series: chart.spec.valueDimensions.map((valueDimension) => {
      const valueKey = getDimensionEncodeKey(dataset, valueDimension);

      return {
        ...(chart.spec.chartType === "area" ? { areaStyle: {} } : {}),
        ...(chart.spec.smooth ? { smooth: true } : {}),
        ...(chart.spec.stacked ? { stack: "total" } : {}),
        encode: {
          itemName: categoryKey,
          tooltip: [categoryKey, valueKey],
          x: categoryKey,
          y: valueKey,
        },
        name: dataset.dimensions[valueDimension]?.name ?? `Series ${valueDimension + 1}`,
        type: seriesType,
      };
    }),
    tooltip: {
      trigger: chart.spec.chartType === "scatter" ? "item" : "axis",
    },
    xAxis: {
      name: xAxisName,
      nameGap: 28,
      nameLocation: "middle",
      type:
        chart.spec.chartType === "scatter"
          ? toAxisType(categoryDimension)
          : categoryDimension === "time"
            ? "time"
            : "category",
    },
    yAxis: {
      name: yAxisName,
      nameGap: yAxisNameGap,
      nameLocation: "middle",
      nameRotate: 90,
      type: "value",
    },
  };
}

function buildPieChartOption(
  chart: WorkbookPieChart,
  dataset: WorkbookChartPreviewDataset,
): Record<string, unknown> {
  const nameKey = getDimensionEncodeKey(dataset, chart.spec.nameDimension);
  const valueKey = getDimensionEncodeKey(dataset, chart.spec.valueDimension);
  const valueName = getDimensionLabel(dataset, chart.spec.valueDimension);

  return {
    dataset: {
      dimensions: dataset.dimensions,
      source: dataset.source,
      sourceHeader: dataset.sourceHeader,
    },
    legend: {
      show: true,
    },
    series: [
      {
        encode: {
          itemName: nameKey,
          tooltip: [nameKey, valueKey],
          value: valueKey,
        },
        label: {
          formatter: "{b}: {d}%",
        },
        name: valueName,
        type: "pie",
      },
    ],
    tooltip: {
      formatter: "{b}: {c} ({d}%)",
      trigger: "item",
    },
  };
}

function getDimensionLabel(dataset: WorkbookChartPreviewDataset, dimensionIndex: number): string {
  return dataset.dimensions[dimensionIndex]?.name ?? `Dimension ${dimensionIndex + 1}`;
}

function getValueAxisLabel(
  dataset: WorkbookChartPreviewDataset,
  valueDimensions: readonly number[],
): string {
  if (valueDimensions.length === 1) {
    return getDimensionLabel(dataset, valueDimensions[0]);
  }

  return "Value";
}

function getValueAxisNameGap(
  chart: WorkbookCartesianChart,
  dataset: WorkbookChartPreviewDataset,
): number {
  const widestTickLabel = getEstimatedValueAxisTickLabelWidth(chart, dataset);

  return Math.max(
    Y_AXIS_NAME_GAP_MIN,
    Math.ceil(widestTickLabel + Y_AXIS_TICK_LABEL_MARGIN + Y_AXIS_NAME_PADDING),
  );
}

function getEstimatedValueAxisTickLabelWidth(
  chart: WorkbookCartesianChart,
  dataset: WorkbookChartPreviewDataset,
): number {
  const extent = getValueAxisExtent(chart, dataset);
  const candidates = [0, extent.min, extent.max]
    .filter((value) => Number.isFinite(value))
    .map((value) => formatValueAxisTickLabel(value));

  return candidates.reduce(
    (width, label) => Math.max(width, estimateAxisLabelWidth(label, AXIS_LABEL_FONT_SIZE)),
    0,
  );
}

function getValueAxisExtent(
  chart: WorkbookCartesianChart,
  dataset: WorkbookChartPreviewDataset,
): { max: number; min: number } {
  let min = 0;
  let max = 0;

  for (const row of getPreviewDataRows(dataset)) {
    if (chart.spec.stacked) {
      let negativeTotal = 0;
      let positiveTotal = 0;

      for (const dimensionIndex of chart.spec.valueDimensions) {
        const value = getNumericPreviewValue(row, dimensionIndex);

        if (value === undefined) {
          continue;
        }

        if (value < 0) {
          negativeTotal += value;
        } else {
          positiveTotal += value;
        }
      }

      min = Math.min(min, negativeTotal);
      max = Math.max(max, positiveTotal);
      continue;
    }

    for (const dimensionIndex of chart.spec.valueDimensions) {
      const value = getNumericPreviewValue(row, dimensionIndex);

      if (value === undefined) {
        continue;
      }

      min = Math.min(min, value);
      max = Math.max(max, value);
    }
  }

  return { max, min };
}

function getPreviewDataRows(
  dataset: WorkbookChartPreviewDataset,
): Array<Array<string | number | null>> {
  return dataset.sourceHeader ? dataset.source.slice(1) : dataset.source;
}

function getNumericPreviewValue(
  row: ReadonlyArray<string | number | null>,
  dimensionIndex: number,
): number | undefined {
  const value = row[dimensionIndex];

  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function formatValueAxisTickLabel(value: number): string {
  if (Math.abs(value) >= 1 || value === 0) {
    return new Intl.NumberFormat("en-US", {
      maximumFractionDigits: 2,
    }).format(value);
  }

  return new Intl.NumberFormat("en-US", {
    maximumSignificantDigits: 3,
  }).format(value);
}

function estimateAxisLabelWidth(label: string, fontSize: number): number {
  let width = 0;

  for (const character of label) {
    width += getAxisLabelCharacterWidth(character, fontSize);
  }

  return width;
}

function getAxisLabelCharacterWidth(character: string, fontSize: number): number {
  if (/[0-9]/.test(character)) {
    return fontSize * 0.56;
  }

  if (character === "," || character === ".") {
    return fontSize * 0.28;
  }

  if (character === "-" || character === "+") {
    return fontSize * 0.35;
  }

  return fontSize * 0.5;
}

function getDimensionEncodeKey(
  dataset: WorkbookChartPreviewDataset,
  dimensionIndex: number,
): string | number {
  return dataset.dimensions[dimensionIndex]?.name ?? dimensionIndex;
}

function toAxisType(dimensionType: WorkbookChartDimensionType): "category" | "time" | "value" {
  if (dimensionType === "number") {
    return "value";
  }

  if (dimensionType === "time") {
    return "time";
  }

  return "category";
}

function isCartesianChart(chart: WorkbookChart): chart is WorkbookCartesianChart {
  return chart.spec.family === "cartesian";
}
