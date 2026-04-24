import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import { applyWorkbookTransaction, createWorkbookState } from "./workbook-core";
import { WorkbookController } from "./workbook-controller";
import { serializeWorkbookDocument } from "./workbook-document";

test("WorkbookController exposes raw range reads separately from display-range and cell-data reads", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["1", "2", "=A1+B1"],
          ["text", "", "=A2+1"],
        ],
      },
    ],
  });

  const rawRange = controller.getSheetRange({
    columnCount: 3,
    rowCount: 2,
    startColumn: 0,
    startRow: 0,
  });
  const displayRange = controller.getSheetDisplayRange({
    columnCount: 3,
    rowCount: 2,
    startColumn: 0,
    startRow: 0,
  });
  const formulaCell = controller.getCellData({
    columnIndex: 2,
    rowIndex: 0,
  });
  const valueErrorCell = controller.getCellData({
    columnIndex: 2,
    rowIndex: 1,
  });

  assert.deepEqual(rawRange.values, [
    ["1", "2", "=A1+B1"],
    ["text", "", "=A2+1"],
  ]);
  assert.deepEqual(displayRange.values, [
    ["1", "2", "3"],
    ["text", "", "#VALUE!"],
  ]);
  assert.deepEqual(formulaCell, {
    columnIndex: 2,
    display: "3",
    input: "=A1+B1",
    isFormula: true,
    rowIndex: 0,
    sheetId: rawRange.sheetId,
    sheetName: rawRange.sheetName,
  });
  assert.deepEqual(valueErrorCell, {
    columnIndex: 2,
    display: "#VALUE!",
    errorCode: "VALUE",
    input: "=A2+1",
    isFormula: true,
    rowIndex: 1,
    sheetId: rawRange.sheetId,
    sheetName: rawRange.sheetName,
  });
  assert.equal(controller.getSummary().hasUnsavedChanges, true);
});

test("WorkbookController exposes expanded formula compatibility through the same raw and display reads", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["a", "10", "=SUM(B1:B2)", "=IFERROR(1/0,99)"],
          [
            "b",
            "20",
            '=XLOOKUP("b",A1:A2,B1:B2,"nf")',
            '=TEXTJOIN(", ",TRUE,A1:A2)',
          ],
        ],
      },
    ],
  });

  const rawRange = controller.getSheetRange({
    columnCount: 4,
    rowCount: 2,
    startColumn: 0,
    startRow: 0,
  });
  const displayRange = controller.getSheetDisplayRange({
    columnCount: 4,
    rowCount: 2,
    startColumn: 0,
    startRow: 0,
  });
  const lookupCell = controller.getCellData({
    columnIndex: 2,
    rowIndex: 1,
  });

  assert.deepEqual(rawRange.values, [
    ["a", "10", "=SUM(B1:B2)", "=IFERROR(1/0,99)"],
    ["b", "20", '=XLOOKUP("b",A1:A2,B1:B2,"nf")', '=TEXTJOIN(", ",TRUE,A1:A2)'],
  ]);
  assert.deepEqual(displayRange.values, [
    ["a", "10", "30", "99"],
    ["b", "20", "20", "a, b"],
  ]);
  assert.deepEqual(lookupCell, {
    columnIndex: 2,
    display: "20",
    input: '=XLOOKUP("b",A1:A2,B1:B2,"nf")',
    isFormula: true,
    rowIndex: 1,
    sheetId: rawRange.sheetId,
    sheetName: rawRange.sheetName,
  });
});

test("WorkbookController keeps CSV export on raw input strings even when formulas are present", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["1", "=A1+2"],
          ["", "=B1*2"],
        ],
      },
    ],
  });

  assert.equal(controller.getSheetCsv(), "1,=A1+2\r\n,=B1*2");
  assert.deepEqual(
    controller.getSheetDisplayRange({
      columnCount: 2,
      rowCount: 2,
      startColumn: 0,
      startRow: 0,
    }).values,
    [
      ["1", "3"],
      ["", "6"],
    ],
  );
});

test("WorkbookController supports raw-vs-display range copy plus explicit cut, paste, and clear helpers", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [["1", "2", "=A1+B1"]],
      },
    ],
  });

  const rawCopy = controller.copyRange({
    columnCount: 3,
    mode: "raw",
    rowCount: 1,
    startColumn: 0,
    startRow: 0,
  });
  const displayCopy = controller.copyRange({
    columnCount: 3,
    mode: "display",
    rowCount: 1,
    startColumn: 0,
    startRow: 0,
  });

  assert.equal(rawCopy.text, "1\t2\t=A1+B1");
  assert.deepEqual(rawCopy.values, [["1", "2", "=A1+B1"]]);
  assert.equal(displayCopy.text, "1\t2\t3");
  assert.deepEqual(displayCopy.values, [["1", "2", "3"]]);

  const cutResult = controller.cutRange({
    columnCount: 3,
    mode: "display",
    rowCount: 1,
    startColumn: 0,
    startRow: 0,
  });

  assert.equal(cutResult.changed, true);
  assert.equal(cutResult.text, "1\t2\t3");
  assert.equal(cutResult.clipboard.rawText, "1\t2\t=A1+B1");
  assert.equal(cutResult.clipboard.displayText, "1\t2\t3");
  assert.deepEqual(cutResult.clipboard.rawValues, [["1", "2", "=A1+B1"]]);
  assert.deepEqual(cutResult.clipboard.displayValues, [["1", "2", "3"]]);
  assert.deepEqual(
    controller.getSheetRange({
      columnCount: 3,
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    }).values,
    [["", "", ""]],
  );

  controller.pasteRange({
    startColumn: 0,
    startRow: 1,
    text: cutResult.clipboard.displayText,
  });

  assert.deepEqual(
    controller.getSheetRange({
      columnCount: 3,
      rowCount: 2,
      startColumn: 0,
      startRow: 0,
    }).values,
    [
      ["", "", ""],
      ["1", "2", "3"],
    ],
  );

  controller.clearRange({
    columnCount: 2,
    rowCount: 1,
    startColumn: 1,
    startRow: 1,
  });

  assert.deepEqual(
    controller.getSheetRange({
      columnCount: 3,
      rowCount: 2,
      startColumn: 0,
      startRow: 0,
    }).values,
    [
      ["", "", ""],
      ["1", "", ""],
    ],
  );
});

test("WorkbookController exposes persisted chart reads and normalized preview data", async () => {
  let workbook = applyWorkbookTransaction(createWorkbookState(), {
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["Quarter", "Revenue", "Cost"],
          ["Q1", "10", "4"],
          ["Q2", "15", "7"],
          ["Q3", "20", "8"],
        ],
      },
      {
        activate: false,
        columnCount: 6,
        name: "Metrics",
        rowCount: 6,
        type: "addSheet",
      },
    ],
  }).state;
  const metricsSheet = workbook.sheets[1];

  workbook = applyWorkbookTransaction(workbook, {
    operations: [
      {
        sheetId: metricsSheet.id,
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["Metric", "Q1", "Q2", "Q3"],
          ["Revenue", "10", "=10/0", "30"],
          ["Cost", "4", "5", "6"],
        ],
      },
    ],
  }).state;
  workbook.charts = [
    {
      id: "chart-1",
      layout: {
        height: 260,
        offsetX: 0,
        offsetY: 0,
        startColumn: 4,
        startRow: 0,
        width: 420,
        zIndex: 0,
      },
      name: "Quarterly Revenue",
      sheetId: workbook.sheets[0].id,
      spec: {
        categoryDimension: 0,
        chartType: "bar",
        family: "cartesian",
        source: {
          range: {
            columnCount: 3,
            rowCount: 4,
            sheetId: workbook.sheets[0].id,
            startColumn: 0,
            startRow: 0,
          },
          seriesLayoutBy: "column",
          sourceHeader: true,
        },
        valueDimensions: [1],
      },
    },
    {
      id: "chart-2",
      layout: {
        height: 260,
        offsetX: 0,
        offsetY: 0,
        startColumn: 4,
        startRow: 8,
        width: 420,
        zIndex: 1,
      },
      name: "Broken Chart",
      sheetId: workbook.sheets[0].id,
      spec: {
        categoryDimension: 0,
        chartType: "line",
        family: "cartesian",
        source: {
          range: {
            columnCount: 0,
            rowCount: 0,
            sheetId: workbook.sheets[0].id,
            startColumn: 0,
            startRow: 0,
          },
          seriesLayoutBy: "column",
          sourceHeader: true,
        },
        valueDimensions: [1],
      },
    },
    {
      id: "chart-3",
      layout: {
        height: 260,
        offsetX: 0,
        offsetY: 0,
        startColumn: 4,
        startRow: 0,
        width: 420,
        zIndex: 2,
      },
      name: "Metrics By Quarter",
      sheetId: metricsSheet.id,
      spec: {
        categoryDimension: 0,
        chartType: "line",
        family: "cartesian",
        source: {
          range: {
            columnCount: 4,
            rowCount: 3,
            sheetId: metricsSheet.id,
            startColumn: 0,
            startRow: 0,
          },
          seriesLayoutBy: "row",
          sourceHeader: true,
        },
        valueDimensions: [1, 2],
      },
    },
  ];
  workbook.nextChartNumber = 4;

  const tempDirectory = await fs.mkdtemp(
    path.join(os.tmpdir(), "spready-controller-charts-"),
  );
  const filePath = path.join(tempDirectory, "charts.spready");

  try {
    await fs.writeFile(filePath, serializeWorkbookDocument(workbook), "utf8");

    const controller = new WorkbookController();
    const openResult = await controller.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath,
    });
    const activeSheetCharts = controller.getSheetCharts();
    const metricsCharts = controller.getSheetCharts(metricsSheet.id);
    const chartResult = controller.getChart("chart-1");
    const preview = controller.getChartPreview("chart-1");
    const sheetPreviews = controller.getSheetChartPreviews();
    const invalidPreview = controller.getChartPreview("chart-2");
    const rowLayoutPreview = controller.getChartPreview("chart-3");

    assert.equal(openResult.summary.charts.length, 3);
    assert.deepEqual(
      openResult.summary.charts.map((chart) => ({
        id: chart.id,
        status: chart.status,
      })),
      [
        {
          id: "chart-1",
          status: "ok",
        },
        {
          id: "chart-2",
          status: "invalid",
        },
        {
          id: "chart-3",
          status: "ok",
        },
      ],
    );
    assert.equal(activeSheetCharts.sheetId, workbook.sheets[0].id);
    assert.deepEqual(
      activeSheetCharts.charts.map((chart) => chart.id),
      ["chart-1", "chart-2"],
    );
    assert.deepEqual(
      sheetPreviews.previews.map((sheetPreview) => sheetPreview.chart.id),
      ["chart-1", "chart-2"],
    );
    assert.equal(sheetPreviews.sheetId, workbook.sheets[0].id);
    assert.equal(metricsCharts.sheetId, metricsSheet.id);
    assert.deepEqual(
      metricsCharts.charts.map((chart) => chart.id),
      ["chart-3"],
    );
    assert.deepEqual(chartResult, {
      chart: workbook.charts[0],
      status: "ok",
      validationIssues: [],
    });
    assert.deepEqual(preview.dataset, {
      dimensions: [
        {
          name: "Quarter",
          type: "ordinal",
        },
        {
          name: "Revenue",
          type: "number",
        },
        {
          name: "Cost",
          type: "number",
        },
      ],
      seriesLayoutBy: "column",
      source: [
        ["Quarter", "Revenue", "Cost"],
        ["Q1", 10, 4],
        ["Q2", 15, 7],
        ["Q3", 20, 8],
      ],
      sourceHeader: true,
    });
    assert.deepEqual(preview.option, {
      dataset: {
        dimensions: preview.dataset.dimensions,
        source: preview.dataset.source,
        sourceHeader: true,
      },
      grid: {
        bottom: 18,
        containLabel: true,
        left: 56,
        right: 16,
        top: 16,
      },
      legend: {
        show: false,
      },
      series: [
        {
          encode: {
            itemName: "Quarter",
            tooltip: ["Quarter", "Revenue"],
            x: "Quarter",
            y: "Revenue",
          },
          name: "Revenue",
          type: "bar",
        },
      ],
      tooltip: {
        trigger: "axis",
      },
      xAxis: {
        name: "Quarter",
        nameGap: 28,
        nameLocation: "middle",
        type: "category",
      },
      yAxis: {
        name: "Revenue",
        nameGap: 42,
        nameLocation: "middle",
        nameRotate: 90,
        type: "value",
      },
    });
    assert.deepEqual(preview.validationIssues, []);
    assert.deepEqual(preview.warnings, []);
    assert.deepEqual(
      invalidPreview.validationIssues.map((issue) => issue.code),
      ["EMPTY_RANGE", "INVALID_DIMENSION"],
    );
    assert.equal(invalidPreview.status, "invalid");
    assert.deepEqual(invalidPreview.dataset, {
      dimensions: [],
      seriesLayoutBy: "column",
      source: [],
      sourceHeader: false,
    });
    assert.deepEqual(invalidPreview.option, {
      series: [],
      title: {
        text: "Broken Chart",
      },
    });
    assert.deepEqual(rowLayoutPreview.dataset, {
      dimensions: [
        {
          name: "Metric",
          type: "ordinal",
        },
        {
          name: "Revenue",
          type: "number",
        },
        {
          name: "Cost",
          type: "number",
        },
      ],
      seriesLayoutBy: "column",
      source: [
        ["Metric", "Revenue", "Cost"],
        ["Q1", 10, 4],
        ["Q2", null, 5],
        ["Q3", 30, 6],
      ],
      sourceHeader: true,
    });
    assert.deepEqual(rowLayoutPreview.option, {
      dataset: {
        dimensions: rowLayoutPreview.dataset.dimensions,
        source: rowLayoutPreview.dataset.source,
        sourceHeader: true,
      },
      grid: {
        bottom: 18,
        containLabel: true,
        left: 56,
        right: 16,
        top: 16,
      },
      legend: {
        show: true,
      },
      series: [
        {
          encode: {
            itemName: "Metric",
            tooltip: ["Metric", "Revenue"],
            x: "Metric",
            y: "Revenue",
          },
          name: "Revenue",
          type: "line",
        },
        {
          encode: {
            itemName: "Metric",
            tooltip: ["Metric", "Cost"],
            x: "Metric",
            y: "Cost",
          },
          name: "Cost",
          type: "line",
        },
      ],
      tooltip: {
        trigger: "axis",
      },
      xAxis: {
        name: "Metric",
        nameGap: 28,
        nameLocation: "middle",
        type: "category",
      },
      yAxis: {
        name: "Value",
        nameGap: 42,
        nameLocation: "middle",
        nameRotate: 90,
        type: "value",
      },
    });
    assert.deepEqual(rowLayoutPreview.validationIssues, []);
    assert.deepEqual(rowLayoutPreview.warnings, [
      "Chart preview skipped one or more formula errors by converting them to null values.",
    ]);
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("WorkbookController labels pie chart previews with slice and value labels", () => {
  const controller = new WorkbookController();
  const sheetId = controller.getSummary().activeSheetId;

  controller.applyTransaction({
    operations: [
      {
        sheetId,
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["Quarter", "Revenue"],
          ["Q1", "10"],
          ["Q2", "12"],
          ["Q3", "15"],
        ],
      },
      {
        chartId: "chart-pie-labels",
        name: "Revenue Share",
        spec: {
          chartType: "pie",
          family: "pie",
          nameDimension: 0,
          source: {
            range: {
              columnCount: 2,
              rowCount: 4,
              sheetId,
              startColumn: 0,
              startRow: 0,
            },
            seriesLayoutBy: "column",
            sourceHeader: true,
          },
          valueDimension: 1,
        },
        type: "addChart",
      },
    ],
  });

  const preview = controller.getChartPreview("chart-pie-labels");

  assert.deepEqual(preview.option, {
    dataset: {
      dimensions: preview.dataset.dimensions,
      source: preview.dataset.source,
      sourceHeader: true,
    },
    legend: {
      show: true,
    },
    series: [
      {
        encode: {
          itemName: "Quarter",
          tooltip: ["Quarter", "Revenue"],
          value: "Revenue",
        },
        label: {
          formatter: "{b}: {d}%",
        },
        name: "Revenue",
        type: "pie",
      },
    ],
    tooltip: {
      formatter: "{b}: {c} ({d}%)",
      trigger: "item",
    },
  });
});

test("WorkbookController creates charts through the simplified chart request contract", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["Month", "Revenue", "Cost"],
          ["Jan", "120", "80"],
          ["Feb", "150", "94"],
        ],
      },
    ],
  });

  const dryRun = controller.createChart({
    chartType: "line",
    dryRun: true,
    name: "Revenue Trend",
  });

  assert.equal(dryRun.changed, true);
  assert.equal(dryRun.version, controller.getSummary().version);
  assert.equal(dryRun.chart.id, "chart-1");
  assert.deepEqual(controller.getSheetCharts().charts, []);

  const result = controller.createChart({
    chartType: "line",
    name: "Revenue Trend",
  });

  assert.equal(result.changed, true);
  assert.equal(result.chart.id, "chart-1");
  assert.equal(result.chart.name, "Revenue Trend");
  assert.deepEqual(controller.getChart("chart-1").chart.spec, {
    categoryDimension: 0,
    chartType: "line",
    family: "cartesian",
    smooth: false,
    source: {
      range: {
        columnCount: 3,
        rowCount: 3,
        sheetId: controller.getSummary().activeSheetId,
        startColumn: 0,
        startRow: 0,
      },
      seriesLayoutBy: "column",
      sourceHeader: true,
    },
    valueDimensions: [1, 2],
  });
});

test("WorkbookController applies chart lifecycle transactions through the shared write path", () => {
  const controller = new WorkbookController();

  const addResult = controller.applyTransaction({
    operations: [
      {
        spec: {
          categoryDimension: 0,
          chartType: "bar",
          family: "cartesian",
          source: {
            range: {
              columnCount: 2,
              rowCount: 4,
              sheetId: controller.getSummary().activeSheetId,
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
    ],
  });

  assert.equal(addResult.changed, true);
  assert.deepEqual(
    controller.getSheetCharts().charts.map((chart) => ({
      id: chart.id,
      name: chart.name,
      sheetId: chart.sheetId,
      type: chart.spec.chartType,
    })),
    [
      {
        id: "chart-1",
        name: "Chart 1",
        sheetId: controller.getSummary().activeSheetId,
        type: "bar",
      },
    ],
  );

  controller.applyTransaction({
    operations: [
      {
        chartId: "chart-1",
        name: "  Trend  ",
        type: "renameChart",
      },
      {
        chartId: "chart-1",
        spec: {
          categoryDimension: 0,
          chartType: "line",
          family: "cartesian",
          source: {
            range: {
              columnCount: 2,
              rowCount: 4,
              sheetId: controller.getSummary().activeSheetId,
              startColumn: 1,
              startRow: 2,
            },
            seriesLayoutBy: "column",
            sourceHeader: true,
          },
          smooth: true,
          valueDimensions: [1],
        },
        type: "setChartSpec",
      },
    ],
  });

  controller.applyTransaction({
    operations: [
      {
        chartId: "chart-1",
        layout: {
          height: 320,
          offsetX: 16,
          offsetY: 9,
          startColumn: 2,
          startRow: 3,
          width: 520,
          zIndex: 5,
        },
        type: "setChartLayout",
      },
    ],
  });

  assert.deepEqual(controller.getChart("chart-1"), {
    chart: {
      id: "chart-1",
      layout: {
        height: 320,
        offsetX: 16,
        offsetY: 9,
        startColumn: 2,
        startRow: 3,
        width: 520,
        zIndex: 5,
      },
      name: "Trend",
      sheetId: controller.getSummary().activeSheetId,
      spec: {
        categoryDimension: 0,
        chartType: "line",
        family: "cartesian",
        smooth: true,
        source: {
          range: {
            columnCount: 2,
            rowCount: 4,
            sheetId: controller.getSummary().activeSheetId,
            startColumn: 1,
            startRow: 2,
          },
          seriesLayoutBy: "column",
          sourceHeader: true,
        },
        valueDimensions: [1],
      },
    },
    status: "ok",
    validationIssues: [],
  });

  controller.applyTransaction({
    operations: [
      {
        chartId: "chart-1",
        type: "deleteChart",
      },
    ],
  });

  assert.deepEqual(controller.getSheetCharts().charts, []);
  assert.throws(() => controller.getChart("chart-1"), /was not found/);
});

test("WorkbookController rejects stale applyTransaction requests with expectedVersion", () => {
  const controller = new WorkbookController();
  const initialVersion = controller.getSummary().version;

  controller.applyTransaction({
    operations: [
      {
        columnIndex: 0,
        rowIndex: 0,
        type: "setCell",
        value: "draft",
      },
    ],
  });

  assert.throws(
    () =>
      controller.applyTransaction({
        expectedVersion: initialVersion,
        operations: [
          {
            columnIndex: 1,
            rowIndex: 0,
            type: "setCell",
            value: "stale",
          },
        ],
      }),
    /Expected workbook version 0, but current version is 1\./,
  );
});

test("WorkbookController saves and opens native workbook files", async () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        startColumn: 0,
        startRow: 0,
        type: "setRange",
        values: [
          ["1", "2", "=A1+B1"],
          ["North", "980", ""],
        ],
      },
      {
        sourceFilePath: "C:\\imports\\north.csv",
        type: "setSheetSourceFile",
      },
      {
        activate: true,
        columnCount: 2,
        name: "Notes",
        rowCount: 2,
        type: "addSheet",
      },
      {
        columnIndex: 0,
        rowIndex: 0,
        type: "setCell",
        value: "Saved",
      },
    ],
  });

  const tempDirectory = await fs.mkdtemp(
    path.join(os.tmpdir(), "spready-controller-"),
  );

  try {
    assert.equal(controller.getSummary().hasUnsavedChanges, true);

    const saveResult = await controller.saveWorkbookFile({
      filePath: path.join(tempDirectory, "budget"),
    });

    assert.equal(saveResult.changed, true);
    assert.equal(saveResult.summary.documentFilePath, saveResult.filePath);
    assert.equal(saveResult.summary.hasUnsavedChanges, false);
    assert.match(saveResult.filePath, /\.spready$/);
    assert.match(
      await fs.readFile(saveResult.filePath, "utf8"),
      /"format": "spready-workbook"/,
    );

    controller.applyTransaction({
      operations: [
        {
          columnIndex: 0,
          rowIndex: 0,
          type: "setCell",
          value: "Changed",
        },
      ],
    });

    assert.equal(controller.getSummary().hasUnsavedChanges, true);

    await assert.rejects(
      () =>
        controller.openWorkbookFile({
          filePath: saveResult.filePath,
        }),
      /discardUnsavedChanges: true/,
    );

    const openResult = await controller.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath: saveResult.filePath,
    });

    assert.equal(openResult.changed, true);
    assert.equal(openResult.summary.documentFilePath, saveResult.filePath);
    assert.equal(openResult.summary.hasUnsavedChanges, false);
    assert.equal(
      controller.getSheetDisplayRange({
        columnCount: 3,
        rowCount: 2,
        sheetId: openResult.summary.sheets[0].id,
        startColumn: 0,
        startRow: 0,
      }).values[0][2],
      "3",
    );
    assert.equal(
      controller.getCellData({
        columnIndex: 0,
        rowIndex: 0,
      }).input,
      "Saved",
    );

    const secondSaveResult = await controller.saveWorkbookFile({
      filePath: saveResult.filePath,
    });

    assert.equal(secondSaveResult.changed, false);
    assert.equal(secondSaveResult.summary.hasUnsavedChanges, false);
  } finally {
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("WorkbookController creates a new workbook and guards unsaved replacement", () => {
  const controller = new WorkbookController();

  controller.applyTransaction({
    operations: [
      {
        columnIndex: 0,
        rowIndex: 0,
        type: "setCell",
        value: "draft",
      },
    ],
  });

  assert.equal(controller.getSummary().hasUnsavedChanges, true);

  assert.throws(
    () => controller.createNewWorkbook(),
    /discardUnsavedChanges: true/,
  );

  const resetResult = controller.createNewWorkbook({
    discardUnsavedChanges: true,
  });

  assert.equal(resetResult.changed, true);
  assert.equal(resetResult.summary.documentFilePath, undefined);
  assert.equal(resetResult.summary.hasUnsavedChanges, false);
  assert.equal(resetResult.summary.sheets.length, 1);
  assert.equal(
    controller.getCellData({
      columnIndex: 0,
      rowIndex: 0,
    }).input,
    "",
  );
});
