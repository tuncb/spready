import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildChartEditorOperations,
  createChartEditorFormState,
  createChartEditorFormStateFromChart,
  getChartEditorSheetId,
  getChartEditorValidationIssues,
  type ChartEditorWindowRequest,
} from "./chart-editor-state";
import type { WorkbookChart, WorkbookSummary, UsedRangeResult } from "./workbook-core";

const summary: WorkbookSummary = {
  activeSheetId: "sheet-1",
  activeSheetName: "Sheet 1",
  charts: [],
  hasUnsavedChanges: false,
  sheets: [
    {
      columnCount: 10,
      id: "sheet-1",
      name: "Sheet 1",
      rowCount: 20,
    },
  ],
  version: 3,
};

const usedRange: UsedRangeResult = {
  columnCount: 2,
  rowCount: 6,
  sheetId: "sheet-1",
  sheetName: "Sheet 1",
  startColumn: 0,
  startRow: 0,
};

const chart: WorkbookChart = {
  id: "chart-1",
  name: "Revenue Scatter",
  sheetId: "sheet-1",
  spec: {
    categoryDimension: 0,
    chartType: "scatter",
    family: "cartesian",
    source: {
      range: {
        columnCount: 2,
        rowCount: 6,
        sheetId: "sheet-1",
        startColumn: 0,
        startRow: 0,
      },
      seriesLayoutBy: "column",
      sourceHeader: true,
    },
    valueDimensions: [1],
  },
};

test("chart editor state creates valid default create-form operations", () => {
  const request: ChartEditorWindowRequest = {
    mode: "create",
    sheetId: "sheet-1",
  };
  const formState = createChartEditorFormState(request, summary, usedRange);

  assert.equal(formState.chartType, "scatter");
  assert.equal(formState.name, "New Scatter Chart");
  assert.deepEqual(buildChartEditorOperations(request, "sheet-1", formState), [
    {
      name: "New Scatter Chart",
      spec: {
        categoryDimension: 0,
        chartType: "scatter",
        family: "cartesian",
        source: {
          range: {
            columnCount: 2,
            rowCount: 6,
            sheetId: "sheet-1",
            startColumn: 0,
            startRow: 0,
          },
          seriesLayoutBy: "column",
          sourceHeader: true,
        },
        valueDimensions: [1],
      },
      type: "addChart",
    },
  ]);
  assert.deepEqual(
    getChartEditorValidationIssues(formState, "sheet-1", summary),
    [],
  );
});

test("chart editor state builds edit operations from persisted chart info", () => {
  const request: ChartEditorWindowRequest = {
    chartId: "chart-1",
    mode: "edit",
  };
  const formState = createChartEditorFormStateFromChart(chart);

  assert.equal(getChartEditorSheetId(request, summary, chart), "sheet-1");
  assert.deepEqual(buildChartEditorOperations(request, "sheet-1", formState), [
    {
      chartId: "chart-1",
      name: "Revenue Scatter",
      type: "renameChart",
    },
    {
      chartId: "chart-1",
      spec: chart.spec,
      type: "setChartSpec",
    },
  ]);
});

test("chart editor state reports invalid settings before submit", () => {
  const request: ChartEditorWindowRequest = {
    mode: "create",
    sheetId: "sheet-1",
  };
  const formState = createChartEditorFormState(request, summary, usedRange);

  formState.name = "  ";
  formState.valueDimensions = "";

  assert.deepEqual(
    getChartEditorValidationIssues(formState, "sheet-1", summary).map(
      (issue) => issue.message,
    ),
    ["Chart name is required."],
  );
});
