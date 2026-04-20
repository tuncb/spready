import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import { applyWorkbookTransaction, createWorkbookState } from "./workbook-core";
import { SpreadyControlClient } from "./control-client";
import { SpreadyControlServer } from "./control-server";
import { WorkbookController } from "./workbook-controller";
import { serializeWorkbookDocument } from "./workbook-document";

test("SpreadyControlServer exposes formula-aware reads over TCP", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    await client.applyTransaction({
      operations: [
        {
          startColumn: 0,
          startRow: 0,
          type: "setRange",
          values: [["4", "5", "=A1+B1"]],
        },
      ],
    });

    const methods = await client.call<string[]>("listMethods");
    const rawCopy = await client.copyRange({
      columnCount: 3,
      mode: "raw",
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });
    const displayCopy = await client.copyRange({
      columnCount: 3,
      mode: "display",
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });
    const cutResult = await client.cutRange({
      columnCount: 3,
      mode: "display",
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });
    const displayRange = await client.getSheetDisplayRange({
      columnCount: 3,
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });
    const cellData = await client.getCellData({
      columnIndex: 2,
      rowIndex: 0,
    });

    assert.ok(methods.includes("getCellData"));
    assert.ok(methods.includes("getSheetDisplayRange"));
    assert.ok(methods.includes("copyRange"));
    assert.ok(methods.includes("cutRange"));
    assert.ok(methods.includes("pasteRange"));
    assert.ok(methods.includes("clearRange"));
    assert.equal(rawCopy.text, "4\t5\t=A1+B1");
    assert.equal(displayCopy.text, "4\t5\t9");
    assert.equal(cutResult.text, "4\t5\t9");
    assert.equal(cutResult.clipboard.rawText, "4\t5\t=A1+B1");
    assert.equal(cutResult.clipboard.displayText, "4\t5\t9");
    assert.deepEqual(displayRange.values, [["", "", ""]]);
    assert.deepEqual(cellData, {
      columnIndex: 2,
      display: "",
      input: "",
      isFormula: false,
      rowIndex: 0,
      sheetId: displayRange.sheetId,
      sheetName: displayRange.sheetName,
    });
  } finally {
    await client.close();
    await server.stop();
  }
});

test("SpreadyControlServer exposes expanded formula compatibility over TCP", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    await client.applyTransaction({
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

    const displayRange = await client.getSheetDisplayRange({
      columnCount: 4,
      rowCount: 2,
      startColumn: 0,
      startRow: 0,
    });
    const cellData = await client.getCellData({
      columnIndex: 2,
      rowIndex: 1,
    });

    assert.deepEqual(displayRange.values, [
      ["a", "10", "30", "99"],
      ["b", "20", "20", "a, b"],
    ]);
    assert.deepEqual(cellData, {
      columnIndex: 2,
      display: "20",
      input: '=XLOOKUP("b",A1:A2,B1:B2,"nf")',
      isFormula: true,
      rowIndex: 1,
      sheetId: displayRange.sheetId,
      sheetName: displayRange.sheetName,
    });
  } finally {
    await client.close();
    await server.stop();
  }
});

test("SpreadyControlServer exposes chart reads and preview data over TCP", async () => {
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
    path.join(os.tmpdir(), "spready-tcp-"),
  );
  const filePath = path.join(tempDirectory, "charts.spready");
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await fs.writeFile(filePath, serializeWorkbookDocument(workbook), "utf8");
  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    const openResult = await client.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath,
    });
    const methods = await client.call<string[]>("listMethods");
    const activeSheetCharts = await client.getSheetCharts();
    const metricsCharts = await client.getSheetCharts(metricsSheet.id);
    const chartResult = await client.getChart("chart-1");
    const invalidPreview = await client.getChartPreview("chart-2");
    const rowLayoutPreview = await client.getChartPreview("chart-3");

    assert.ok(methods.includes("getChart"));
    assert.ok(methods.includes("getChartPreview"));
    assert.ok(methods.includes("getSheetCharts"));
    assert.equal(openResult.summary.charts.length, 3);
    assert.deepEqual(
      activeSheetCharts.charts.map((chart) => chart.id),
      ["chart-1", "chart-2"],
    );
    assert.deepEqual(
      metricsCharts.charts.map((chart) => chart.id),
      ["chart-3"],
    );
    assert.equal(chartResult.status, "ok");
    assert.deepEqual(chartResult.validationIssues, []);
    assert.deepEqual(
      invalidPreview.validationIssues.map((issue) => issue.code),
      ["EMPTY_RANGE", "INVALID_DIMENSION"],
    );
    assert.equal(invalidPreview.status, "invalid");
    assert.deepEqual(rowLayoutPreview.dataset.source, [
      ["Metric", "Revenue", "Cost"],
      ["Q1", 10, 4],
      ["Q2", null, 5],
      ["Q3", 30, 6],
    ]);
    assert.deepEqual(rowLayoutPreview.warnings, [
      "Chart preview skipped one or more formula errors by converting them to null values.",
    ]);
    assert.deepEqual(rowLayoutPreview.option, {
      dataset: {
        dimensions: rowLayoutPreview.dataset.dimensions,
        source: rowLayoutPreview.dataset.source,
        sourceHeader: true,
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
      title: {
        text: "Metrics By Quarter",
      },
      tooltip: {
        trigger: "axis",
      },
      xAxis: {
        type: "category",
      },
      yAxis: {
        type: "value",
      },
    });
  } finally {
    await client.close();
    await server.stop();
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("SpreadyControlServer saves and opens native workbook files over TCP", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);
  const tempDirectory = await fs.mkdtemp(
    path.join(os.tmpdir(), "spready-tcp-"),
  );

  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    await client.applyTransaction({
      operations: [
        {
          startColumn: 0,
          startRow: 0,
          type: "setRange",
          values: [["4", "5", "=A1+B1"]],
        },
      ],
    });

    const filePath = path.join(tempDirectory, "numbers.spready");
    const saveResult = await client.saveWorkbookFile({ filePath });

    assert.equal(saveResult.changed, true);
    assert.equal(saveResult.summary.documentFilePath, filePath);
    assert.equal(saveResult.summary.hasUnsavedChanges, false);
    assert.ok(
      (await client.call<string[]>("listMethods")).includes("openWorkbookFile"),
    );
    assert.ok(
      (await client.call<string[]>("listMethods")).includes("saveWorkbookFile"),
    );

    await client.applyTransaction({
      operations: [
        {
          columnIndex: 0,
          rowIndex: 0,
          type: "setCell",
          value: "99",
        },
      ],
    });

    await assert.rejects(
      () => client.openWorkbookFile({ filePath }),
      /discardUnsavedChanges: true/,
    );

    const openResult = await client.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath,
    });
    const displayRange = await client.getSheetDisplayRange({
      columnCount: 3,
      rowCount: 1,
      startColumn: 0,
      startRow: 0,
    });

    assert.equal(openResult.summary.documentFilePath, filePath);
    assert.equal(openResult.summary.hasUnsavedChanges, false);
    assert.deepEqual(displayRange.values, [["4", "5", "9"]]);
  } finally {
    await client.close();
    await server.stop();
    await fs.rm(tempDirectory, { force: true, recursive: true });
  }
});

test("SpreadyControlServer creates a new workbook over TCP with unsaved-change guard", async () => {
  const controller = new WorkbookController();
  const server = new SpreadyControlServer(controller, "127.0.0.1", 0);

  await server.start();

  const controlInfo = server.getInfo();
  const client = new SpreadyControlClient({
    host: controlInfo.host,
    port: controlInfo.port,
    source: "argv",
  });

  try {
    await client.connect();

    await client.applyTransaction({
      operations: [
        {
          columnIndex: 0,
          rowIndex: 0,
          type: "setCell",
          value: "draft",
        },
      ],
    });

    await assert.rejects(
      () => client.createNewWorkbook(),
      /discardUnsavedChanges: true/,
    );

    const result = await client.createNewWorkbook({
      discardUnsavedChanges: true,
    });

    assert.equal(result.changed, true);
    assert.equal(result.summary.documentFilePath, undefined);
    assert.equal(result.summary.hasUnsavedChanges, false);
    assert.equal(
      (
        await client.getCellData({
          columnIndex: 0,
          rowIndex: 0,
        })
      ).input,
      "",
    );
  } finally {
    await client.close();
    await server.stop();
  }
});
