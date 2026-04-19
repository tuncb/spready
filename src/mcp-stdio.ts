#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  SubscribeRequestSchema,
  UnsubscribeRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import * as z from "zod/v4";

import { CONTROL_DISCOVERY_FILE_PATH } from "./control-discovery";
import { resolveControlTarget, SpreadyControlClient } from "./control-client";

const workbookSummarySchema = z.object({
  activeSheetId: z.string(),
  activeSheetName: z.string(),
  documentFilePath: z.string().optional(),
  hasUnsavedChanges: z.boolean(),
  sheets: z.array(
    z.object({
      columnCount: z.int().min(1),
      id: z.string(),
      name: z.string(),
      rowCount: z.int().min(1),
      sourceFilePath: z.string().optional(),
    }),
  ),
  version: z.int().min(0),
});

const usedRangeSchema = z.object({
  columnCount: z.int().min(0),
  rowCount: z.int().min(0),
  sheetId: z.string(),
  sheetName: z.string(),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
});

const sheetRangeSchema = z.object({
  columnCount: z.int().min(0),
  rowCount: z.int().min(0),
  sheetId: z.string(),
  sheetName: z.string(),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
  values: z.array(z.array(z.string())),
});

const sheetDisplayRangeSchema = z.object({
  columnCount: z.int().min(0),
  rowCount: z.int().min(0),
  sheetId: z.string(),
  sheetName: z.string(),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
  values: z.array(z.array(z.string())),
});

const cellDataSchema = z.object({
  columnIndex: z.int().min(0),
  display: z.string(),
  errorCode: z.enum(["PARSE", "REF", "DIV0", "VALUE", "CYCLE"]).optional(),
  input: z.string(),
  isFormula: z.boolean(),
  rowIndex: z.int().min(0),
  sheetId: z.string(),
  sheetName: z.string(),
});

const clipboardRangeModeSchema = z.enum(["display", "raw"]);

const copyRangeResultSchema = z.object({
  columnCount: z.int().min(0),
  mode: clipboardRangeModeSchema,
  rowCount: z.int().min(0),
  sheetId: z.string(),
  sheetName: z.string(),
  startColumn: z.int().min(0),
  startRow: z.int().min(0),
  text: z.string(),
  values: z.array(z.array(z.string())),
});

const transactionOperationSchema = z.discriminatedUnion("type", [
  z.object({
    activate: z.boolean().optional(),
    columnCount: z.int().min(1).optional(),
    name: z.string().min(1).optional(),
    rowCount: z.int().min(1).optional(),
    sheetId: z.string().min(1).optional(),
    type: z.literal("addSheet"),
  }),
  z.object({
    columnCount: z.int().min(0),
    rowCount: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    startColumn: z.int().min(0),
    startRow: z.int().min(0),
    type: z.literal("clearRange"),
  }),
  z.object({
    columnIndex: z.int().min(0),
    count: z.int().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal("deleteColumns"),
  }),
  z.object({
    count: z.int().min(1),
    rowIndex: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    type: z.literal("deleteRows"),
  }),
  z.object({
    nextActiveSheetId: z.string().min(1).optional(),
    sheetId: z.string().min(1),
    type: z.literal("deleteSheet"),
  }),
  z.object({
    columnIndex: z.int().min(0),
    count: z.int().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal("insertColumns"),
  }),
  z.object({
    count: z.int().min(1),
    rowIndex: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    type: z.literal("insertRows"),
  }),
  z.object({
    name: z.string().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal("renameSheet"),
  }),
  z.object({
    name: z.string().min(1).optional(),
    rows: z.array(z.array(z.string())),
    sheetId: z.string().min(1).optional(),
    sourceFilePath: z.string().min(1).optional(),
    type: z.literal("replaceSheet"),
  }),
  z.object({
    content: z.string(),
    name: z.string().min(1).optional(),
    sheetId: z.string().min(1).optional(),
    sourceFilePath: z.string().min(1).optional(),
    type: z.literal("replaceSheetFromCsv"),
  }),
  z.object({
    columnCount: z.int().min(1),
    rowCount: z.int().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal("resizeSheet"),
  }),
  z.object({
    sheetId: z.string().min(1),
    type: z.literal("setActiveSheet"),
  }),
  z.object({
    sheetId: z.string().min(1).optional(),
    sourceFilePath: z.string().optional(),
    type: z.literal("setSheetSourceFile"),
  }),
  z.object({
    columnIndex: z.int().min(0),
    rowIndex: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    type: z.literal("setCell"),
    value: z.string(),
  }),
  z.object({
    sheetId: z.string().min(1).optional(),
    startColumn: z.int().min(0),
    startRow: z.int().min(0),
    type: z.literal("setRange"),
    values: z.array(z.array(z.string())),
  }),
]);

const applyTransactionResultSchema = z.object({
  changed: z.boolean(),
  summary: workbookSummarySchema,
  version: z.int().min(0),
});

const csvFileOperationResultSchema = z.object({
  changed: z.boolean(),
  filePath: z.string(),
  summary: workbookSummarySchema,
  version: z.int().min(0),
});

const workbookFileOperationResultSchema = z.object({
  changed: z.boolean(),
  filePath: z.string(),
  summary: workbookSummarySchema,
  version: z.int().min(0),
});

const WORKBOOK_SUMMARY_RESOURCE_URI = "spready://workbook/summary";
const SERVER_GUIDE_RESOURCE_URI = "spready://guide";
const WORKBOOK_TASK_PROMPT_NAME = "spready_workbook_task";

const transactionOperations = [
  {
    type: "addSheet",
    description:
      "Add a new sheet, optionally naming it, sizing it, and making it active.",
  },
  {
    type: "clearRange",
    description: "Clear a rectangular range without resizing the sheet.",
  },
  {
    type: "deleteColumns",
    description:
      "Delete one or more columns starting at a zero-based column index.",
  },
  {
    type: "deleteRows",
    description: "Delete one or more rows starting at a zero-based row index.",
  },
  {
    type: "deleteSheet",
    description:
      "Delete a sheet by id and optionally choose which sheet becomes active next.",
  },
  {
    type: "insertColumns",
    description:
      "Insert one or more blank columns at a zero-based column index.",
  },
  {
    type: "insertRows",
    description: "Insert one or more blank rows at a zero-based row index.",
  },
  {
    type: "renameSheet",
    description: "Rename an existing sheet.",
  },
  {
    type: "replaceSheet",
    description: "Replace an entire sheet from an in-memory 2D string array.",
  },
  {
    type: "replaceSheetFromCsv",
    description: "Replace an entire sheet from CSV content.",
  },
  {
    type: "resizeSheet",
    description: "Set the row and column counts for a sheet.",
  },
  {
    type: "setActiveSheet",
    description: "Make a specific sheet active.",
  },
  {
    type: "setSheetSourceFile",
    description: "Attach or clear the source file path metadata for a sheet.",
  },
  {
    type: "setCell",
    description: "Write a single string value to one cell.",
  },
  {
    type: "setRange",
    description:
      "Write a rectangular 2D string array starting at a zero-based row and column.",
  },
] as const;

const guideResourceSchema = z.object({
  prompt: z.object({
    description: z.string(),
    name: z.string(),
  }),
  resources: z.array(
    z.object({
      description: z.string(),
      mimeType: z.string(),
      name: z.string(),
      uri: z.string(),
    }),
  ),
  startupRequirement: z.string(),
  tools: z.array(
    z.object({
      defaultsToActiveSheet: z.boolean(),
      description: z.string(),
      name: z.string(),
      readOnly: z.boolean(),
    }),
  ),
  transactionOperations: z.array(
    z.object({
      description: z.string(),
      type: z.string(),
    }),
  ),
  usageConventions: z.array(z.string()),
  workflow: z.array(z.string()),
});

const guideResource = {
  prompt: {
    description:
      "Prompt template for planning or executing one workbook task with Spready.",
    name: WORKBOOK_TASK_PROMPT_NAME,
  },
  resources: [
    {
      description: "Current workbook summary for the connected Spready app.",
      mimeType: "application/json",
      name: "workbook-summary",
      uri: WORKBOOK_SUMMARY_RESOURCE_URI,
    },
    {
      description:
        "Usage guide for third-party harnesses integrating with Spready.",
      mimeType: "text/markdown",
      name: "server-guide",
      uri: SERVER_GUIDE_RESOURCE_URI,
    },
  ],
  startupRequirement:
    "The Spready desktop app must already be running before this MCP wrapper can connect.",
  tools: [
    {
      defaultsToActiveSheet: false,
      description:
        "Return workbook metadata including active sheet, version, and sheet sizes.",
      name: "get_workbook_summary",
      readOnly: true,
    },
    {
      defaultsToActiveSheet: false,
      description:
        "Create a new blank workbook and replace the in-app workbook state.",
      name: "create_new_workbook",
      readOnly: false,
    },
    {
      defaultsToActiveSheet: false,
      description:
        "Open a native Spready workbook file and replace the in-app workbook state.",
      name: "open_workbook_file",
      readOnly: false,
    },
    {
      defaultsToActiveSheet: false,
      description:
        "Save the current workbook as a native Spready workbook file.",
      name: "save_workbook_file",
      readOnly: false,
    },
    {
      defaultsToActiveSheet: true,
      description: "Return the used range bounds for a sheet.",
      name: "get_used_range",
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description:
        "Read one cell with both raw input and evaluated display output.",
      name: "get_cell_data",
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description:
        "Read a rectangular range of evaluated display values from a sheet.",
      name: "get_sheet_display_range",
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description: "Read a rectangular range from a sheet.",
      name: "get_sheet_range",
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description:
        "Copy one rectangular range as tab-delimited text, using raw cell input or displayed values.",
      name: "copy_range",
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description: "Return the sheet as CSV text trimmed to its used range.",
      name: "get_sheet_csv",
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description:
        "Paste tab-delimited text or explicit values into a rectangular range starting at one cell.",
      name: "paste_range",
      readOnly: false,
    },
    {
      defaultsToActiveSheet: true,
      description: "Clear a rectangular range without resizing the sheet.",
      name: "clear_range",
      readOnly: false,
    },
    {
      defaultsToActiveSheet: true,
      description: "Apply one atomic batch of workbook mutations.",
      name: "apply_transaction",
      readOnly: false,
    },
    {
      defaultsToActiveSheet: true,
      description:
        "Import a local CSV file into a sheet and update its source file metadata.",
      name: "import_csv_file",
      readOnly: false,
    },
    {
      defaultsToActiveSheet: true,
      description:
        "Export one sheet as CSV to a local file and update its source file metadata.",
      name: "export_csv_file",
      readOnly: false,
    },
  ],
  transactionOperations: transactionOperations.map((operation) => ({
    description: operation.description,
    type: operation.type,
  })),
  usageConventions: [
    "Rows and columns are zero-based.",
    "Use create_new_workbook for a fresh blank workbook.",
    "Use open_workbook_file and save_workbook_file for full multi-sheet workbook persistence.",
    "Check hasUnsavedChanges in get_workbook_summary before replacing the current workbook.",
    "Read tools default to the active sheet when sheetId is omitted.",
    "Use get_sheet_range for raw workbook input and get_sheet_display_range for evaluated grid values.",
    "Use copy_range when you need a tab-delimited clipboard-style payload for one explicit rectangular range.",
    "Use get_cell_data when you need one cell's raw formula text plus its evaluated display result.",
    "Many transaction operations also default to the active sheet when sheetId is omitted.",
    "Use import_csv_file and export_csv_file only for single-sheet CSV interchange.",
    "Use get_workbook_summary before large edits so you know which sheet ids and sizes exist.",
    "Use get_used_range or get_sheet_range instead of reading an entire large sheet.",
    "Use paste_range and clear_range for explicit clipboard-like range edits without relying on UI selection state.",
    "Prefer one apply_transaction call with batched operations over repeated single-cell writes.",
    "Use dryRun on apply_transaction to validate risky changes before mutating the workbook.",
    "CSV file paths are resolved on the same machine running the Spready desktop app and MCP wrapper.",
    `Subscribe to ${WORKBOOK_SUMMARY_RESOURCE_URI} if your client supports live workbook summary updates.`,
  ],
  workflow: [
    "Create a new workbook with create_new_workbook when the task should start from a blank workbook.",
    "Open an existing workbook with open_workbook_file when the task starts from a .spready document.",
    "Inspect the workbook with get_workbook_summary.",
    "Narrow the target area with get_used_range or get_sheet_range.",
    "Read only the rows or columns needed for the task.",
    "Validate planned edits with apply_transaction dryRun when the change is risky or destructive.",
    "Apply the final mutation in one batched apply_transaction call.",
    "Save the finished workbook with save_workbook_file when you need a durable workbook document.",
  ],
} satisfies z.infer<typeof guideResourceSchema>;

const guideMarkdown = `# Spready MCP Guide

## Startup requirement

${guideResource.startupRequirement}

## Recommended workflow

${guideResource.workflow
  .map((step, index) => `${index + 1}. ${step}`)
  .join("\n")}

## Usage conventions

- ${guideResource.usageConventions.join("\n- ")}

## Tools

- get_workbook_summary: Return workbook metadata including active sheet, version, and sheet sizes.
- create_new_workbook: Create a new blank workbook and replace the in-app workbook state.
- open_workbook_file: Open a native Spready workbook file and replace the in-app workbook state.
- save_workbook_file: Save the current workbook as a native Spready workbook file.
- get_used_range: Return the used range bounds for a sheet. Omitting sheetId uses the active sheet.
- get_cell_data: Return one cell's raw input plus its evaluated display value. Omitting sheetId uses the active sheet.
- get_sheet_display_range: Read one rectangular range of evaluated display values. Prefer this for formula-aware grid views.
- get_sheet_range: Read one rectangular range. Prefer this over loading a large sheet.
- copy_range: Return one rectangular range plus tab-delimited text using raw input or displayed values.
- get_sheet_csv: Return trimmed CSV for one sheet. Omitting sheetId uses the active sheet.
- paste_range: Paste tab-delimited text or explicit values into the sheet starting at one cell. Omitting sheetId uses the active sheet.
- clear_range: Clear one explicit rectangular range without resizing the sheet. Omitting sheetId uses the active sheet.
- apply_transaction: Apply one atomic batch of workbook mutations. Supports dryRun.
- import_csv_file: Load a local CSV file into a sheet. Omitting sheetId uses the active sheet.
- export_csv_file: Save one sheet as a local CSV file. Omitting sheetId uses the active sheet.

## Transaction operation types

- ${transactionOperations
  .map((operation) => `${operation.type}: ${operation.description}`)
  .join("\n- ")}

## Example read request

\`\`\`json
{
  "startRow": 0,
  "startColumn": 0,
  "rowCount": 20,
  "columnCount": 5
}
\`\`\`

## Example dry-run transaction

\`\`\`json
{
  "dryRun": true,
  "operations": [
    {
      "type": "setRange",
      "startRow": 0,
      "startColumn": 0,
      "values": [
        ["Region", "Revenue"],
        ["North", "1200"],
        ["South", "980"]
      ]
    }
  ]
}
\`\`\`

## Prompt

- ${WORKBOOK_TASK_PROMPT_NAME}: Prompt template for planning or executing one workbook task with Spready.
`;

const describeCapabilitiesResultSchema = z.object({
  overview: z.string(),
  prompt: z.object({
    description: z.string(),
    name: z.string(),
  }),
  resources: guideResourceSchema.shape.resources,
  startupRequirement: z.string(),
  tools: z.array(
    z.object({
      defaultsToActiveSheet: z.boolean(),
      description: z.string(),
      name: z.string(),
      readOnly: z.boolean(),
      useWhen: z.string(),
    }),
  ),
  transactionOperations: guideResourceSchema.shape.transactionOperations,
  usageConventions: z.array(z.string()),
  workflow: z.array(z.string()),
});

function createTextResult<Result extends object>(payload: Result) {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(payload, null, 2),
      },
    ],
    structuredContent: payload as Record<string, unknown>,
  };
}

function getArgumentValue(name: string): string | undefined {
  const index = process.argv.indexOf(name);

  if (index < 0) {
    return undefined;
  }

  return process.argv[index + 1];
}

function parsePort(portValue?: string): number | undefined {
  if (!portValue) {
    return undefined;
  }

  const port = Number.parseInt(portValue, 10);

  if (Number.isNaN(port)) {
    throw new Error(`Invalid port "${portValue}".`);
  }

  return port;
}

async function main() {
  const target = await resolveControlTarget({
    host: getArgumentValue("--host"),
    port: parsePort(getArgumentValue("--port")),
  });
  const controlClient = new SpreadyControlClient(target);

  try {
    await controlClient.connect();
  } catch (error) {
    const detail =
      error instanceof Error ? error.message : "unknown connection error";

    throw new Error(
      `Could not connect to the Spready control server at tcp://${target.host}:${target.port}. ` +
        `Start the Electron app first or set SPREADY_CONTROL_HOST/SPREADY_CONTROL_PORT. ` +
        `Discovery file: ${CONTROL_DISCOVERY_FILE_PATH}. ${detail}`,
    );
  }

  const server = new McpServer(
    {
      name: "spready-stdio",
      version: "0.0.2",
    },
    {
      capabilities: {
        logging: {},
        resources: {
          subscribe: true,
        },
      },
      instructions:
        "Spready requires the desktop app to already be running. Start with describe_capabilities or read spready://guide, use open_workbook_file and save_workbook_file for native workbook documents, inspect with get_workbook_summary before large edits, use zero-based indexes, use get_sheet_range for raw input and get_sheet_display_range for evaluated grid values, and prefer apply_transaction with batched operations plus dryRun for risky changes.",
    },
  );
  const subscribedResourceUris = new Set<string>();
  const knownResourceUris = new Set([
    SERVER_GUIDE_RESOURCE_URI,
    WORKBOOK_SUMMARY_RESOURCE_URI,
  ]);

  controlClient.on("workbookChanged", async () => {
    if (!subscribedResourceUris.has(WORKBOOK_SUMMARY_RESOURCE_URI)) {
      return;
    }

    try {
      await server.server.sendResourceUpdated({
        uri: WORKBOOK_SUMMARY_RESOURCE_URI,
      });
    } catch {
      // Ignore notification failures when the client does not consume resource updates.
    }
  });

  server.server.setRequestHandler(SubscribeRequestSchema, async (request) => {
    const { uri } = request.params;

    if (!knownResourceUris.has(uri)) {
      throw new Error(`Unknown resource "${uri}".`);
    }

    subscribedResourceUris.add(uri);
    return {};
  });

  server.server.setRequestHandler(UnsubscribeRequestSchema, async (request) => {
    subscribedResourceUris.delete(request.params.uri);
    return {};
  });

  server.registerResource(
    "workbook-summary",
    WORKBOOK_SUMMARY_RESOURCE_URI,
    {
      description: "Current workbook summary for the connected Spready app.",
      mimeType: "application/json",
      title: "Workbook Summary",
    },
    async () => {
      const summary = await controlClient.getWorkbookSummary();

      return {
        contents: [
          {
            mimeType: "application/json",
            text: JSON.stringify(summary, null, 2),
            uri: WORKBOOK_SUMMARY_RESOURCE_URI,
          },
        ],
      };
    },
  );

  server.registerResource(
    "server-guide",
    SERVER_GUIDE_RESOURCE_URI,
    {
      description:
        "Usage guide for third-party harnesses integrating with Spready.",
      mimeType: "text/markdown",
      title: "Spready MCP Guide",
    },
    async () => {
      return {
        contents: [
          {
            mimeType: "text/markdown",
            text: guideMarkdown,
            uri: SERVER_GUIDE_RESOURCE_URI,
          },
        ],
      };
    },
  );

  server.registerTool(
    "get_workbook_summary",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Return workbook metadata including active sheet, version, and sheet sizes.",
      outputSchema: workbookSummarySchema,
    },
    async () => createTextResult(await controlClient.getWorkbookSummary()),
  );

  server.registerTool(
    "create_new_workbook",
    {
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Create a new blank workbook and replace the in-app workbook state.",
      inputSchema: z.object({
        discardUnsavedChanges: z
          .boolean()
          .optional()
          .describe(
            "Set to true to replace the current workbook even when it has unsaved changes.",
          ),
      }),
      outputSchema: applyTransactionResultSchema,
    },
    async (args) => createTextResult(await controlClient.createNewWorkbook(args)),
  );

  server.registerTool(
    "open_workbook_file",
    {
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Open a native Spready workbook file and replace the in-app workbook state.",
      inputSchema: z.object({
        discardUnsavedChanges: z
          .boolean()
          .optional()
          .describe(
            "Set to true to replace the current workbook even when it has unsaved changes.",
          ),
        filePath: z
          .string()
          .min(1)
          .describe(
            "Path to a .spready workbook file on the machine running Spready.",
          ),
      }),
      outputSchema: workbookFileOperationResultSchema,
    },
    async (args) => createTextResult(await controlClient.openWorkbookFile(args)),
  );

  server.registerTool(
    "save_workbook_file",
    {
      annotations: {
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Save the current workbook as a native Spready workbook file.",
      inputSchema: z.object({
        filePath: z
          .string()
          .min(1)
          .describe(
            "Destination path for a .spready workbook file on the machine running Spready.",
          ),
      }),
      outputSchema: workbookFileOperationResultSchema,
    },
    async (args) => createTextResult(await controlClient.saveWorkbookFile(args)),
  );

  server.registerTool(
    "describe_capabilities",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Describe the Spready MCP server, recommended workflow, usage conventions, and supported transaction operations.",
      outputSchema: describeCapabilitiesResultSchema,
    },
    async () =>
      createTextResult({
        overview:
          "Spready is a spreadsheet-focused MCP server. Read the workbook first, then apply batched transactions for edits.",
        prompt: guideResource.prompt,
        resources: guideResource.resources,
        startupRequirement: guideResource.startupRequirement,
        tools: [
          {
            defaultsToActiveSheet: false,
            description:
              "Return workbook metadata including active sheet, version, and sheet sizes.",
            name: "get_workbook_summary",
            readOnly: true,
            useWhen:
              "Always use this first before exploring or editing a workbook, and inspect hasUnsavedChanges before replacing it.",
          },
          {
            defaultsToActiveSheet: false,
            description:
              "Create a new blank workbook and replace the in-app workbook state.",
            name: "create_new_workbook",
            readOnly: false,
            useWhen:
              "Use this when the task should start from a blank workbook. Pass discardUnsavedChanges only after you have explicitly decided to replace unsaved work.",
          },
          {
            defaultsToActiveSheet: false,
            description:
              "Open a native Spready workbook file and replace the in-app workbook state.",
            name: "open_workbook_file",
            readOnly: false,
            useWhen:
              "Use this when the task starts from a saved .spready workbook document. Pass discardUnsavedChanges only after you have explicitly decided to replace unsaved work.",
          },
          {
            defaultsToActiveSheet: false,
            description:
              "Save the current workbook as a native Spready workbook file.",
            name: "save_workbook_file",
            readOnly: false,
            useWhen:
              "Use this when the result should persist as a full multi-sheet workbook document.",
          },
          {
            defaultsToActiveSheet: true,
            description: "Return the used range bounds for a sheet.",
            name: "get_used_range",
            readOnly: true,
            useWhen:
              "Use this to find the populated area before reading a range or exporting CSV.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Return one cell with both raw input and evaluated display output.",
            name: "get_cell_data",
            readOnly: true,
            useWhen:
              "Use this when the task needs a formula cell's raw text and computed result.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Read a rectangular range of evaluated display values from a sheet.",
            name: "get_sheet_display_range",
            readOnly: true,
            useWhen:
              "Use this for formula-aware grid reads where displayed results matter.",
          },
          {
            defaultsToActiveSheet: true,
            description: "Read a rectangular range from a sheet.",
            name: "get_sheet_range",
            readOnly: true,
            useWhen:
              "Use this for targeted inspection of raw workbook input strings.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Copy one rectangular range as tab-delimited text using raw input or displayed values.",
            name: "copy_range",
            readOnly: true,
            useWhen:
              "Use this when a task needs explicit clipboard-style range output without relying on UI selection state.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Return the sheet as CSV text trimmed to its used range.",
            name: "get_sheet_csv",
            readOnly: true,
            useWhen:
              "Use this when you need the full used range in one text payload.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Paste tab-delimited text or explicit values into the sheet starting at one cell.",
            name: "paste_range",
            readOnly: false,
            useWhen:
              "Use this for explicit range pastes when the input is already tabular text or a 2D string array.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Clear a rectangular range without resizing the sheet.",
            name: "clear_range",
            readOnly: false,
            useWhen:
              "Use this for delete-style range clearing without constructing a full apply_transaction request.",
          },
          {
            defaultsToActiveSheet: true,
            description: "Apply one atomic batch of workbook mutations.",
            name: "apply_transaction",
            readOnly: false,
            useWhen:
              "Use this for all writes, preferably in one batched request with dryRun first.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Import a local CSV file into a sheet and update its source file metadata.",
            name: "import_csv_file",
            readOnly: false,
            useWhen:
              "Use this when your task starts from a CSV file that already exists on disk.",
          },
          {
            defaultsToActiveSheet: true,
            description:
              "Export one sheet as CSV to a local file and update its source file metadata.",
            name: "export_csv_file",
            readOnly: false,
            useWhen:
              "Use this when the final result should be written to a CSV file on disk.",
          },
        ],
        transactionOperations: guideResource.transactionOperations,
        usageConventions: guideResource.usageConventions,
        workflow: guideResource.workflow,
      }),
  );

  server.registerTool(
    "get_used_range",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Return the used range bounds for a sheet. Omit sheetId to use the active sheet.",
      inputSchema: z.object({
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
      }),
      outputSchema: usedRangeSchema,
    },
    async ({ sheetId }) =>
      createTextResult(await controlClient.getUsedRange(sheetId)),
  );

  server.registerTool(
    "get_cell_data",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Return one cell with both the raw stored input and the evaluated display value. Omit sheetId to use the active sheet.",
      inputSchema: z.object({
        columnIndex: z
          .int()
          .min(0)
          .describe("Zero-based column index of the cell to read."),
        rowIndex: z
          .int()
          .min(0)
          .describe("Zero-based row index of the cell to read."),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
      }),
      outputSchema: cellDataSchema,
    },
    async (args) => createTextResult(await controlClient.getCellData(args)),
  );

  server.registerTool(
    "get_sheet_display_range",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Read a rectangular range of evaluated display values from a sheet. Use this for formula-aware grid reads.",
      inputSchema: z.object({
        columnCount: z.int().min(1).describe("Number of columns to read."),
        rowCount: z.int().min(1).describe("Number of rows to read."),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
        startColumn: z.int().min(0).describe("Zero-based start column."),
        startRow: z.int().min(0).describe("Zero-based start row."),
      }),
      outputSchema: sheetDisplayRangeSchema,
    },
    async (args) =>
      createTextResult(await controlClient.getSheetDisplayRange(args)),
  );

  server.registerTool(
    "get_sheet_range",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Read a rectangular range from a sheet. Use this instead of reading an entire large sheet.",
      inputSchema: z.object({
        columnCount: z.int().min(1).describe("Number of columns to read."),
        rowCount: z.int().min(1).describe("Number of rows to read."),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
        startColumn: z.int().min(0).describe("Zero-based start column."),
        startRow: z.int().min(0).describe("Zero-based start row."),
      }),
      outputSchema: sheetRangeSchema,
    },
    async (args) => createTextResult(await controlClient.getSheetRange(args)),
  );

  server.registerTool(
    "copy_range",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Copy one rectangular range as tab-delimited text using raw cell input or displayed values.",
      inputSchema: z.object({
        columnCount: z.int().min(1).describe("Number of columns to copy."),
        mode: clipboardRangeModeSchema
          .optional()
          .describe(
            'Copy mode. Use "raw" to preserve formulas, or "display" to flatten them to visible values.',
          ),
        rowCount: z.int().min(1).describe("Number of rows to copy."),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
        startColumn: z.int().min(0).describe("Zero-based start column."),
        startRow: z.int().min(0).describe("Zero-based start row."),
      }),
      outputSchema: copyRangeResultSchema,
    },
    async (args) => createTextResult(await controlClient.copyRange(args)),
  );

  server.registerTool(
    "get_sheet_csv",
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        "Return the sheet as CSV text trimmed to its used range. Omit sheetId to use the active sheet.",
      inputSchema: z.object({
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
      }),
      outputSchema: z.object({
        csv: z.string(),
      }),
    },
    async ({ sheetId }) => {
      const csv = await controlClient.getSheetCsv(sheetId);

      return createTextResult({ csv });
    },
  );

  server.registerTool(
    "import_csv_file",
    {
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Import a local CSV file into a sheet and update that sheet source file metadata.",
      inputSchema: z.object({
        filePath: z
          .string()
          .min(1)
          .describe("Path to a local CSV file on the machine running Spready."),
        name: z
          .string()
          .min(1)
          .optional()
          .describe("Optional sheet name override to apply during import."),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
      }),
      outputSchema: csvFileOperationResultSchema,
    },
    async (args) => createTextResult(await controlClient.importCsvFile(args)),
  );

  server.registerTool(
    "export_csv_file",
    {
      annotations: {
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Export one sheet as CSV to a local file and update that sheet source file metadata.",
      inputSchema: z.object({
        filePath: z
          .string()
          .min(1)
          .describe("Destination CSV path on the machine running Spready."),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
      }),
      outputSchema: csvFileOperationResultSchema,
    },
    async (args) => createTextResult(await controlClient.exportCsvFile(args)),
  );

  server.registerTool(
    "paste_range",
    {
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Paste tab-delimited text or explicit values into a sheet starting at one cell.",
      inputSchema: z
        .object({
          sheetId: z
            .string()
            .min(1)
            .optional()
            .describe(
              "Optional target sheet id. Defaults to the active sheet.",
            ),
          startColumn: z.int().min(0).describe("Zero-based start column."),
          startRow: z.int().min(0).describe("Zero-based start row."),
          text: z
            .string()
            .optional()
            .describe(
              "Optional tab-delimited text payload to parse into a 2D string matrix.",
            ),
          values: z
            .array(z.array(z.string()))
            .optional()
            .describe("Optional explicit 2D string matrix to paste."),
        })
        .refine(
          (value) => value.text !== undefined || value.values !== undefined,
          {
            message: 'Provide either "text" or "values".',
            path: ["text"],
          },
        ),
      outputSchema: applyTransactionResultSchema,
    },
    async (args) => createTextResult(await controlClient.pasteRange(args)),
  );

  server.registerTool(
    "clear_range",
    {
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Clear a rectangular range without resizing the sheet. Omit sheetId to use the active sheet.",
      inputSchema: z.object({
        columnCount: z.int().min(1).describe("Number of columns to clear."),
        rowCount: z.int().min(1).describe("Number of rows to clear."),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe("Optional target sheet id. Defaults to the active sheet."),
        startColumn: z.int().min(0).describe("Zero-based start column."),
        startRow: z.int().min(0).describe("Zero-based start row."),
      }),
      outputSchema: applyTransactionResultSchema,
    },
    async (args) => createTextResult(await controlClient.clearRange(args)),
  );

  server.registerTool(
    "apply_transaction",
    {
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        "Apply one atomic batch of workbook mutations. Prefer batched operations over repeated single-cell writes.",
      inputSchema: z.object({
        dryRun: z
          .boolean()
          .optional()
          .describe(
            "Validate and simulate the transaction without mutating the workbook.",
          ),
        operations: z
          .array(transactionOperationSchema)
          .min(1)
          .describe("Ordered transaction operations to apply atomically."),
      }),
      outputSchema: applyTransactionResultSchema,
    },
    async (args) =>
      createTextResult(await controlClient.applyTransaction(args)),
  );

  server.registerPrompt(
    WORKBOOK_TASK_PROMPT_NAME,
    {
      description:
        "Template prompt for planning or executing one workbook task with Spready.",
      argsSchema: {
        goal: z.string().min(1).describe("The workbook task to accomplish."),
      },
    },
    async ({ goal }) => {
      return {
        messages: [
          {
            role: "user",
            content: {
              type: "text",
              text:
                `Use the Spready MCP server to accomplish this workbook task: ${goal}\n\n` +
                "Workflow:\n" +
                "- Use create_new_workbook when the task should start from a blank workbook.\n" +
                "- Use open_workbook_file when the task starts from an existing .spready workbook.\n" +
                "- Start with get_workbook_summary.\n" +
                "- If get_workbook_summary reports hasUnsavedChanges, save first or pass discardUnsavedChanges only if losing local changes is intended.\n" +
                "- Use zero-based row and column indexes.\n" +
                "- Use get_sheet_range for raw workbook input and get_sheet_display_range for evaluated grid values.\n" +
                "- Use get_cell_data when one cell's raw formula text and display result both matter.\n" +
                "- Read only the ranges you need with get_used_range, get_sheet_range, or get_sheet_display_range.\n" +
                "- Use apply_transaction for writes, preferably as one batched request.\n" +
                "- Use save_workbook_file when the final result should persist as a native workbook document.\n" +
                "- Use dryRun before risky or destructive mutations.\n" +
                `- If you need server details, call describe_capabilities or read ${SERVER_GUIDE_RESOURCE_URI}.`,
            },
          },
        ],
      };
    },
  );

  const transport = new StdioServerTransport();

  await server.connect(transport);
  console.error(
    `Spready MCP stdio wrapper connected to tcp://${target.host}:${target.port} via ${target.source}`,
  );
}

main().catch((error) => {
  console.error("Spready MCP stdio wrapper failed:", error);
  process.exit(1);
});
