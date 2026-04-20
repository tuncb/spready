import * as z from "zod/v4";

import type {
  WorkbookChartPreview,
  WorkbookChartResult,
  WorkbookSheetChartsResult,
} from "./workbook-core";

const workbookChartTypeSchema = z.enum([
  "bar",
  "line",
  "area",
  "scatter",
  "pie",
]);
const workbookChartStatusSchema = z.enum(["ok", "invalid"]);
const workbookChartSeriesLayoutSchema = z.enum(["column", "row"]);
const workbookChartDimensionTypeSchema = z.enum(["number", "ordinal", "time"]);
const chartValueSchema = z.union([z.string(), z.number(), z.null()]);

const workbookChartRangeSchema = z.object({
  columnCount: z.int().min(0),
  rowCount: z.int().min(0),
  sheetId: z.string(),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
});

const workbookChartSourceSchema = z.object({
  range: workbookChartRangeSchema,
  seriesLayoutBy: workbookChartSeriesLayoutSchema,
  sourceHeader: z.boolean(),
});

const workbookChartCartesianSpecSchema = z.object({
  categoryDimension: z.int().min(0),
  chartType: z.enum(["bar", "line", "area", "scatter"]),
  family: z.literal("cartesian"),
  smooth: z.boolean().optional(),
  source: workbookChartSourceSchema,
  stacked: z.boolean().optional(),
  valueDimensions: z.array(z.int().min(0)),
});

const workbookChartPieSpecSchema = z.object({
  chartType: z.literal("pie"),
  family: z.literal("pie"),
  nameDimension: z.int().min(0),
  source: workbookChartSourceSchema,
  valueDimension: z.int().min(0),
});

export const workbookChartSchema = z.object({
  id: z.string(),
  name: z.string(),
  sheetId: z.string(),
  spec: z.discriminatedUnion("family", [
    workbookChartCartesianSpecSchema,
    workbookChartPieSpecSchema,
  ]),
});

export const workbookChartSummarySchema = z.object({
  chartType: workbookChartTypeSchema,
  id: z.string(),
  name: z.string(),
  sheetId: z.string(),
  status: workbookChartStatusSchema,
});

const workbookChartValidationIssueSchema = z.object({
  code: z.enum([
    "CROSS_SHEET_SOURCE",
    "EMPTY_RANGE",
    "EMPTY_VALUE_DIMENSIONS",
    "INVALID_DIMENSION",
    "INVALID_RANGE_COORDINATE",
    "MISSING_SHEET",
    "OUT_OF_BOUNDS",
    "REPEATED_VALUE_DIMENSION",
  ]),
  message: z.string(),
});

const workbookChartPreviewDimensionSchema = z.object({
  name: z.string(),
  type: workbookChartDimensionTypeSchema,
});

const workbookChartPreviewDatasetSchema = z.object({
  dimensions: z.array(workbookChartPreviewDimensionSchema),
  seriesLayoutBy: workbookChartSeriesLayoutSchema,
  source: z.array(z.array(chartValueSchema)),
  sourceHeader: z.boolean(),
});

export const workbookSheetChartsResultSchema = z.object({
  charts: z.array(workbookChartSchema),
  sheetId: z.string(),
  sheetName: z.string(),
});

export const workbookChartResultSchema = z.object({
  chart: workbookChartSchema,
  status: workbookChartStatusSchema,
  validationIssues: z.array(workbookChartValidationIssueSchema),
});

export const workbookChartPreviewSchema = workbookChartResultSchema.extend({
  dataset: workbookChartPreviewDatasetSchema,
  option: z.record(z.string(), z.unknown()),
  warnings: z.array(z.string()),
});

const optionalSheetIdInputSchema = z.object({
  sheetId: z
    .string()
    .min(1)
    .optional()
    .describe("Optional target sheet id. Defaults to the active sheet."),
});

const chartIdInputSchema = z.object({
  chartId: z.string().min(1).describe("Chart id to read."),
});

export const chartGuideTools = [
  {
    defaultsToActiveSheet: true,
    description: "Return the chart definitions owned by a sheet.",
    name: "get_sheet_charts",
    readOnly: true,
    useWhen:
      "Use this to discover the charts on a sheet before requesting one chart or preview in detail.",
  },
  {
    defaultsToActiveSheet: false,
    description:
      "Return one chart definition plus validation status and issues.",
    name: "get_chart",
    readOnly: true,
    useWhen:
      "Use this when you need one persisted chart contract and want to inspect whether Spready considers it valid.",
  },
  {
    defaultsToActiveSheet: false,
    description:
      "Return one chart's normalized preview dataset, warnings, and derived ECharts option.",
    name: "get_chart_preview",
    readOnly: true,
    useWhen:
      "Use this when you need renderer-ready chart preview data without recreating workbook logic outside Spready.",
  },
] as const;

interface ChartToolRegistrar {
  registerTool(
    name: string,
    config: {
      annotations: {
        openWorldHint: false;
        readOnlyHint: true;
      };
      description: string;
      inputSchema?: z.ZodTypeAny;
      outputSchema: z.ZodTypeAny;
    },
    handler: (
      args: unknown,
      ...extra: unknown[]
    ) => Promise<{
      content: Array<{ text: string; type: "text" }>;
      structuredContent: Record<string, unknown>;
    }>,
  ): void;
}

interface ChartToolControlClient {
  getChart(chartId: string): Promise<WorkbookChartResult>;
  getChartPreview(chartId: string): Promise<WorkbookChartPreview>;
  getSheetCharts(sheetId?: string): Promise<WorkbookSheetChartsResult>;
}

export function registerChartTools(
  server: ChartToolRegistrar,
  controlClient: ChartToolControlClient,
) {
  server.registerTool(
    "get_sheet_charts",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Return the chart definitions owned by a sheet. Omit sheetId to use the active sheet.",
      inputSchema: optionalSheetIdInputSchema,
      outputSchema: workbookSheetChartsResultSchema,
    },
    async (args) => {
      const { sheetId } = optionalSheetIdInputSchema.parse(args);

      return createTextResult(await controlClient.getSheetCharts(sheetId));
    },
  );

  server.registerTool(
    "get_chart",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Return one chart definition together with validation status and issues.",
      inputSchema: chartIdInputSchema,
      outputSchema: workbookChartResultSchema,
    },
    async (args) => {
      const { chartId } = chartIdInputSchema.parse(args);

      return createTextResult(await controlClient.getChart(chartId));
    },
  );

  server.registerTool(
    "get_chart_preview",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Return one chart's normalized preview dataset, preview warnings, and derived ECharts option.",
      inputSchema: chartIdInputSchema,
      outputSchema: workbookChartPreviewSchema,
    },
    async (args) => {
      const { chartId } = chartIdInputSchema.parse(args);

      return createTextResult(await controlClient.getChartPreview(chartId));
    },
  );
}

function createTextResult<Result extends object>(payload: Result) {
  return {
    content: [
      {
        text: JSON.stringify(payload, null, 2),
        type: "text" as const,
      },
    ],
    structuredContent: payload as Record<string, unknown>,
  };
}
