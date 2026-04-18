#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import * as z from 'zod/v4';

import { CONTROL_DISCOVERY_FILE_PATH } from './control-discovery';
import { resolveControlTarget, SpreadyControlClient } from './control-client';

const workbookSummarySchema = z.object({
  activeSheetId: z.string(),
  activeSheetName: z.string(),
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

const transactionOperationSchema = z.discriminatedUnion('type', [
  z.object({
    activate: z.boolean().optional(),
    columnCount: z.int().min(1).optional(),
    name: z.string().min(1).optional(),
    rowCount: z.int().min(1).optional(),
    sheetId: z.string().min(1).optional(),
    type: z.literal('addSheet'),
  }),
  z.object({
    columnCount: z.int().min(0),
    rowCount: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    startColumn: z.int().min(0),
    startRow: z.int().min(0),
    type: z.literal('clearRange'),
  }),
  z.object({
    columnIndex: z.int().min(0),
    count: z.int().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal('deleteColumns'),
  }),
  z.object({
    count: z.int().min(1),
    rowIndex: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    type: z.literal('deleteRows'),
  }),
  z.object({
    nextActiveSheetId: z.string().min(1).optional(),
    sheetId: z.string().min(1),
    type: z.literal('deleteSheet'),
  }),
  z.object({
    columnIndex: z.int().min(0),
    count: z.int().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal('insertColumns'),
  }),
  z.object({
    count: z.int().min(1),
    rowIndex: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    type: z.literal('insertRows'),
  }),
  z.object({
    name: z.string().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal('renameSheet'),
  }),
  z.object({
    name: z.string().min(1).optional(),
    rows: z.array(z.array(z.string())),
    sheetId: z.string().min(1).optional(),
    sourceFilePath: z.string().min(1).optional(),
    type: z.literal('replaceSheet'),
  }),
  z.object({
    content: z.string(),
    name: z.string().min(1).optional(),
    sheetId: z.string().min(1).optional(),
    sourceFilePath: z.string().min(1).optional(),
    type: z.literal('replaceSheetFromCsv'),
  }),
  z.object({
    columnCount: z.int().min(1),
    rowCount: z.int().min(1),
    sheetId: z.string().min(1).optional(),
    type: z.literal('resizeSheet'),
  }),
  z.object({
    sheetId: z.string().min(1),
    type: z.literal('setActiveSheet'),
  }),
  z.object({
    sheetId: z.string().min(1).optional(),
    sourceFilePath: z.string().optional(),
    type: z.literal('setSheetSourceFile'),
  }),
  z.object({
    columnIndex: z.int().min(0),
    rowIndex: z.int().min(0),
    sheetId: z.string().min(1).optional(),
    type: z.literal('setCell'),
    value: z.string(),
  }),
  z.object({
    sheetId: z.string().min(1).optional(),
    startColumn: z.int().min(0),
    startRow: z.int().min(0),
    type: z.literal('setRange'),
    values: z.array(z.array(z.string())),
  }),
]);

const applyTransactionResultSchema = z.object({
  changed: z.boolean(),
  summary: workbookSummarySchema,
  version: z.int().min(0),
});

const WORKBOOK_SUMMARY_RESOURCE_URI = 'spready://workbook/summary';
const SERVER_GUIDE_RESOURCE_URI = 'spready://guide';
const WORKBOOK_TASK_PROMPT_NAME = 'spready_workbook_task';

const transactionOperations = [
  {
    type: 'addSheet',
    description: 'Add a new sheet, optionally naming it, sizing it, and making it active.',
  },
  {
    type: 'clearRange',
    description: 'Clear a rectangular range without resizing the sheet.',
  },
  {
    type: 'deleteColumns',
    description: 'Delete one or more columns starting at a zero-based column index.',
  },
  {
    type: 'deleteRows',
    description: 'Delete one or more rows starting at a zero-based row index.',
  },
  {
    type: 'deleteSheet',
    description: 'Delete a sheet by id and optionally choose which sheet becomes active next.',
  },
  {
    type: 'insertColumns',
    description: 'Insert one or more blank columns at a zero-based column index.',
  },
  {
    type: 'insertRows',
    description: 'Insert one or more blank rows at a zero-based row index.',
  },
  {
    type: 'renameSheet',
    description: 'Rename an existing sheet.',
  },
  {
    type: 'replaceSheet',
    description: 'Replace an entire sheet from an in-memory 2D string array.',
  },
  {
    type: 'replaceSheetFromCsv',
    description: 'Replace an entire sheet from CSV content.',
  },
  {
    type: 'resizeSheet',
    description: 'Set the row and column counts for a sheet.',
  },
  {
    type: 'setActiveSheet',
    description: 'Make a specific sheet active.',
  },
  {
    type: 'setSheetSourceFile',
    description: 'Attach or clear the source file path metadata for a sheet.',
  },
  {
    type: 'setCell',
    description: 'Write a single string value to one cell.',
  },
  {
    type: 'setRange',
    description: 'Write a rectangular 2D string array starting at a zero-based row and column.',
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
    description: 'Prompt template for planning or executing one workbook task with Spready.',
    name: WORKBOOK_TASK_PROMPT_NAME,
  },
  resources: [
    {
      description: 'Current workbook summary for the connected Spready app.',
      mimeType: 'application/json',
      name: 'workbook-summary',
      uri: WORKBOOK_SUMMARY_RESOURCE_URI,
    },
    {
      description: 'Usage guide for third-party harnesses integrating with Spready.',
      mimeType: 'text/markdown',
      name: 'server-guide',
      uri: SERVER_GUIDE_RESOURCE_URI,
    },
  ],
  startupRequirement:
    'The Spready desktop app must already be running before this MCP wrapper can connect.',
  tools: [
    {
      defaultsToActiveSheet: false,
      description: 'Return workbook metadata including active sheet, version, and sheet sizes.',
      name: 'get_workbook_summary',
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description: 'Return the used range bounds for a sheet.',
      name: 'get_used_range',
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description: 'Read a rectangular range from a sheet.',
      name: 'get_sheet_range',
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description: 'Return the sheet as CSV text trimmed to its used range.',
      name: 'get_sheet_csv',
      readOnly: true,
    },
    {
      defaultsToActiveSheet: true,
      description: 'Apply one atomic batch of workbook mutations.',
      name: 'apply_transaction',
      readOnly: false,
    },
  ],
  transactionOperations: transactionOperations.map((operation) => ({
    description: operation.description,
    type: operation.type,
  })),
  usageConventions: [
    'Rows and columns are zero-based.',
    'Read tools default to the active sheet when sheetId is omitted.',
    'Many transaction operations also default to the active sheet when sheetId is omitted.',
    'Use get_workbook_summary before large edits so you know which sheet ids and sizes exist.',
    'Use get_used_range or get_sheet_range instead of reading an entire large sheet.',
    'Prefer one apply_transaction call with batched operations over repeated single-cell writes.',
    'Use dryRun on apply_transaction to validate risky changes before mutating the workbook.',
  ],
  workflow: [
    'Inspect the workbook with get_workbook_summary.',
    'Narrow the target area with get_used_range or get_sheet_range.',
    'Read only the rows or columns needed for the task.',
    'Validate planned edits with apply_transaction dryRun when the change is risky or destructive.',
    'Apply the final mutation in one batched apply_transaction call.',
  ],
} satisfies z.infer<typeof guideResourceSchema>;

const guideMarkdown = `# Spready MCP Guide

## Startup requirement

${guideResource.startupRequirement}

## Recommended workflow

1. ${guideResource.workflow[0]}
2. ${guideResource.workflow[1]}
3. ${guideResource.workflow[2]}
4. ${guideResource.workflow[3]}
5. ${guideResource.workflow[4]}

## Usage conventions

- ${guideResource.usageConventions.join('\n- ')}

## Tools

- get_workbook_summary: Return workbook metadata including active sheet, version, and sheet sizes.
- get_used_range: Return the used range bounds for a sheet. Omitting sheetId uses the active sheet.
- get_sheet_range: Read one rectangular range. Prefer this over loading a large sheet.
- get_sheet_csv: Return trimmed CSV for one sheet. Omitting sheetId uses the active sheet.
- apply_transaction: Apply one atomic batch of workbook mutations. Supports dryRun.

## Transaction operation types

- ${transactionOperations
  .map((operation) => `${operation.type}: ${operation.description}`)
  .join('\n- ')}

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
        type: 'text' as const,
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
    host: getArgumentValue('--host'),
    port: parsePort(getArgumentValue('--port')),
  });
  const controlClient = new SpreadyControlClient(target);

  try {
    await controlClient.connect();
  } catch (error) {
    const detail =
      error instanceof Error ? error.message : 'unknown connection error';

    throw new Error(
      `Could not connect to the Spready control server at tcp://${target.host}:${target.port}. ` +
        `Start the Electron app first or set SPREADY_CONTROL_HOST/SPREADY_CONTROL_PORT. ` +
        `Discovery file: ${CONTROL_DISCOVERY_FILE_PATH}. ${detail}`,
    );
  }

  const server = new McpServer(
    {
      name: 'spready-stdio',
      version: '0.0.2',
    },
    {
      capabilities: {
        logging: {},
        resources: {
          subscribe: false,
        },
      },
      instructions:
        'Spready requires the desktop app to already be running. Start with describe_capabilities or read spready://guide, inspect with get_workbook_summary before large edits, use zero-based indexes, and prefer apply_transaction with batched operations plus dryRun for risky changes.',
    },
  );

  controlClient.on('workbookChanged', async () => {
    try {
      await server.server.sendResourceUpdated({
        uri: WORKBOOK_SUMMARY_RESOURCE_URI,
      });
    } catch {
      // Ignore notification failures when the client does not consume resource updates.
    }
  });

  server.registerResource(
    'workbook-summary',
    WORKBOOK_SUMMARY_RESOURCE_URI,
    {
      description: 'Current workbook summary for the connected Spready app.',
      mimeType: 'application/json',
      title: 'Workbook Summary',
    },
    async () => {
      const summary = await controlClient.getWorkbookSummary();

      return {
        contents: [
          {
            mimeType: 'application/json',
            text: JSON.stringify(summary, null, 2),
            uri: WORKBOOK_SUMMARY_RESOURCE_URI,
          },
        ],
      };
    },
  );

  server.registerResource(
    'server-guide',
    SERVER_GUIDE_RESOURCE_URI,
    {
      description: 'Usage guide for third-party harnesses integrating with Spready.',
      mimeType: 'text/markdown',
      title: 'Spready MCP Guide',
    },
    async () => {
      return {
        contents: [
          {
            mimeType: 'text/markdown',
            text: guideMarkdown,
            uri: SERVER_GUIDE_RESOURCE_URI,
          },
        ],
      };
    },
  );

  server.registerTool(
    'get_workbook_summary',
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description: 'Return workbook metadata including active sheet, version, and sheet sizes.',
      outputSchema: workbookSummarySchema,
    },
    async () => createTextResult(await controlClient.getWorkbookSummary()),
  );

  server.registerTool(
    'describe_capabilities',
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        'Describe the Spready MCP server, recommended workflow, usage conventions, and supported transaction operations.',
      outputSchema: describeCapabilitiesResultSchema,
    },
    async () =>
      createTextResult({
        overview:
          'Spready is a spreadsheet-focused MCP server. Read the workbook first, then apply batched transactions for edits.',
        prompt: guideResource.prompt,
        resources: guideResource.resources,
        startupRequirement: guideResource.startupRequirement,
        tools: [
          {
            defaultsToActiveSheet: false,
            description: 'Return workbook metadata including active sheet, version, and sheet sizes.',
            name: 'get_workbook_summary',
            readOnly: true,
            useWhen: 'Always use this first before exploring or editing a workbook.',
          },
          {
            defaultsToActiveSheet: true,
            description: 'Return the used range bounds for a sheet.',
            name: 'get_used_range',
            readOnly: true,
            useWhen: 'Use this to find the populated area before reading a range or exporting CSV.',
          },
          {
            defaultsToActiveSheet: true,
            description: 'Read a rectangular range from a sheet.',
            name: 'get_sheet_range',
            readOnly: true,
            useWhen: 'Use this for targeted inspection of a subset of rows and columns.',
          },
          {
            defaultsToActiveSheet: true,
            description: 'Return the sheet as CSV text trimmed to its used range.',
            name: 'get_sheet_csv',
            readOnly: true,
            useWhen: 'Use this when you need the full used range in one text payload.',
          },
          {
            defaultsToActiveSheet: true,
            description: 'Apply one atomic batch of workbook mutations.',
            name: 'apply_transaction',
            readOnly: false,
            useWhen: 'Use this for all writes, preferably in one batched request with dryRun first.',
          },
        ],
        transactionOperations: guideResource.transactionOperations,
        usageConventions: guideResource.usageConventions,
        workflow: guideResource.workflow,
      }),
  );

  server.registerTool(
    'get_used_range',
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        'Return the used range bounds for a sheet. Omit sheetId to use the active sheet.',
      inputSchema: z.object({
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe('Optional target sheet id. Defaults to the active sheet.'),
      }),
      outputSchema: usedRangeSchema,
    },
    async ({ sheetId }) => createTextResult(await controlClient.getUsedRange(sheetId)),
  );

  server.registerTool(
    'get_sheet_range',
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        'Read a rectangular range from a sheet. Use this instead of reading an entire large sheet.',
      inputSchema: z.object({
        columnCount: z
          .int()
          .min(1)
          .describe('Number of columns to read.'),
        rowCount: z
          .int()
          .min(1)
          .describe('Number of rows to read.'),
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe('Optional target sheet id. Defaults to the active sheet.'),
        startColumn: z
          .int()
          .min(0)
          .describe('Zero-based start column.'),
        startRow: z
          .int()
          .min(0)
          .describe('Zero-based start row.'),
      }),
      outputSchema: sheetRangeSchema,
    },
    async (args) => createTextResult(await controlClient.getSheetRange(args)),
  );

  server.registerTool(
    'get_sheet_csv',
    {
      annotations: {
        openWorldHint: false,
        readOnlyHint: true,
      },
      description:
        'Return the sheet as CSV text trimmed to its used range. Omit sheetId to use the active sheet.',
      inputSchema: z.object({
        sheetId: z
          .string()
          .min(1)
          .optional()
          .describe('Optional target sheet id. Defaults to the active sheet.'),
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
    'apply_transaction',
    {
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
        readOnlyHint: false,
      },
      description:
        'Apply one atomic batch of workbook mutations. Prefer batched operations over repeated single-cell writes.',
      inputSchema: z.object({
        dryRun: z
          .boolean()
          .optional()
          .describe('Validate and simulate the transaction without mutating the workbook.'),
        operations: z
          .array(transactionOperationSchema)
          .min(1)
          .describe('Ordered transaction operations to apply atomically.'),
      }),
      outputSchema: applyTransactionResultSchema,
    },
    async (args) => createTextResult(await controlClient.applyTransaction(args)),
  );

  server.registerPrompt(
    WORKBOOK_TASK_PROMPT_NAME,
    {
      description: 'Template prompt for planning or executing one workbook task with Spready.',
      argsSchema: {
        goal: z.string().min(1).describe('The workbook task to accomplish.'),
      },
    },
    async ({ goal }) => {
      return {
        messages: [
          {
            role: 'user',
            content: {
              type: 'text',
              text:
                `Use the Spready MCP server to accomplish this workbook task: ${goal}\n\n` +
                'Workflow:\n' +
                '- Start with get_workbook_summary.\n' +
                '- Use zero-based row and column indexes.\n' +
                '- Read only the ranges you need with get_used_range or get_sheet_range.\n' +
                '- Use apply_transaction for writes, preferably as one batched request.\n' +
                '- Use dryRun before risky or destructive mutations.\n' +
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
  console.error('Spready MCP stdio wrapper failed:', error);
  process.exit(1);
});
