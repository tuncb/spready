import assert from "node:assert/strict";
import { test } from "node:test";

import type {
  WorkbookChartPreview,
  WorkbookChartResult,
  WorkbookSheetChartPreviewsResult,
  WorkbookSheetChartsResult,
} from "./workbook-core";
import {
  chartGuideTools,
  registerChartTools,
  workbookChartPreviewSchema,
  workbookChartResultSchema,
  workbookSheetChartPreviewsResultSchema,
  workbookSheetChartsResultSchema,
} from "./mcp-chart-tools";

test("registerChartTools wires thin MCP adapters over the TCP chart methods", async () => {
  const registrations = new Map<
    string,
    {
      config: {
        description: string;
        inputSchema?: { parse: (value: unknown) => unknown };
        outputSchema: { parse: (value: unknown) => unknown };
      };
      handler: (args: unknown) => Promise<{
        content: Array<{ text: string; type: "text" }>;
        structuredContent: Record<string, unknown>;
      }>;
    }
  >();
  const calls: string[] = [];
  const sheetChartsResult: WorkbookSheetChartsResult = {
    charts: [
      {
        id: "chart-1",
        layout: {
          height: 260,
          offsetX: 0,
          offsetY: 0,
          startColumn: 3,
          startRow: 0,
          width: 420,
          zIndex: 0,
        },
        name: "Revenue",
        sheetId: "sheet-1",
        spec: {
          categoryDimension: 0,
          chartType: "bar",
          family: "cartesian" as const,
          source: {
            range: {
              columnCount: 2,
              rowCount: 4,
              sheetId: "sheet-1",
              startColumn: 0,
              startRow: 0,
            },
            seriesLayoutBy: "column" as const,
            sourceHeader: true,
          },
          valueDimensions: [1],
        },
      },
    ],
    sheetId: "sheet-1",
    sheetName: "Sheet 1",
  };
  const chartResult: WorkbookChartResult = {
    chart: sheetChartsResult.charts[0],
    status: "ok" as const,
    validationIssues: [],
  };
  const chartPreviewResult: WorkbookChartPreview = {
    ...chartResult,
    dataset: {
      dimensions: [
        {
          name: "Quarter",
          type: "ordinal" as const,
        },
        {
          name: "Revenue",
          type: "number" as const,
        },
      ],
      seriesLayoutBy: "column" as const,
      source: [
        ["Quarter", "Revenue"],
        ["Q1", 10],
        ["Q2", 12],
      ],
      sourceHeader: true,
    },
    option: {
      series: [
        {
          type: "bar",
        },
      ],
      title: {
        text: "Revenue",
      },
    },
    warnings: [],
  };
  const sheetChartPreviewsResult: WorkbookSheetChartPreviewsResult = {
    previews: [chartPreviewResult],
    sheetId: "sheet-1",
    sheetName: "Sheet 1",
  };

  registerChartTools(
    {
      registerTool(name, config, handler) {
        registrations.set(name, {
          config,
          handler: (args) => handler(args),
        });
      },
    },
    {
      async getChart(chartId) {
        calls.push(`getChart:${chartId}`);
        return chartResult;
      },
      async getChartPreview(chartId) {
        calls.push(`getChartPreview:${chartId}`);
        return chartPreviewResult;
      },
      async getSheetChartPreviews(sheetId) {
        calls.push(`getSheetChartPreviews:${sheetId ?? "active"}`);
        return sheetChartPreviewsResult;
      },
      async getSheetCharts(sheetId) {
        calls.push(`getSheetCharts:${sheetId ?? "active"}`);
        return sheetChartsResult;
      },
    },
  );

  assert.deepEqual(
    [...registrations.keys()],
    ["get_sheet_charts", "get_chart", "get_chart_preview", "get_sheet_chart_previews"],
  );
  assert.deepEqual(
    chartGuideTools.map((tool) => tool.name),
    ["get_sheet_charts", "get_chart", "get_chart_preview", "get_sheet_chart_previews"],
  );

  const sheetChartsRegistration = registrations.get("get_sheet_charts");
  const chartRegistration = registrations.get("get_chart");
  const previewRegistration = registrations.get("get_chart_preview");
  const sheetPreviewsRegistration = registrations.get("get_sheet_chart_previews");

  assert.ok(sheetChartsRegistration);
  assert.ok(chartRegistration);
  assert.ok(previewRegistration);
  assert.ok(sheetPreviewsRegistration);

  sheetChartsRegistration.config.inputSchema?.parse({});
  chartRegistration.config.inputSchema?.parse({ chartId: "chart-1" });
  previewRegistration.config.inputSchema?.parse({ chartId: "chart-1" });
  sheetPreviewsRegistration.config.inputSchema?.parse({});

  const sheetChartsResponse = await sheetChartsRegistration.handler({});
  const chartResponse = await chartRegistration.handler({ chartId: "chart-1" });
  const previewResponse = await previewRegistration.handler({
    chartId: "chart-1",
  });
  const sheetPreviewsResponse = await sheetPreviewsRegistration.handler({});

  assert.deepEqual(calls, [
    "getSheetCharts:active",
    "getChart:chart-1",
    "getChartPreview:chart-1",
    "getSheetChartPreviews:active",
  ]);
  assert.deepEqual(
    workbookSheetChartsResultSchema.parse(sheetChartsResponse.structuredContent),
    sheetChartsResult,
  );
  assert.deepEqual(workbookChartResultSchema.parse(chartResponse.structuredContent), chartResult);
  assert.deepEqual(
    workbookChartPreviewSchema.parse(previewResponse.structuredContent),
    chartPreviewResult,
  );
  assert.deepEqual(
    workbookSheetChartPreviewsResultSchema.parse(sheetPreviewsResponse.structuredContent),
    sheetChartPreviewsResult,
  );
  assert.equal(sheetChartsResponse.content[0]?.type, "text");
  assert.match(sheetChartsResponse.content[0]?.text ?? "", /"sheetId": "sheet-1"/);
  assert.match(chartResponse.content[0]?.text ?? "", /"chart-1"/);
  assert.match(previewResponse.content[0]?.text ?? "", /"Revenue"/);
});
