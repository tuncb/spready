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
        'Use get_workbook_summary before large edits, inspect only the ranges you need, and prefer apply_transaction with batched operations over repeated single-cell writes.',
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
