#!/usr/bin/env node
import { existsSync } from "node:fs";
import { performance } from "node:perf_hooks";
import path from "node:path";

const DEFAULT_WORKBOOK_PATH = "C:\\work\\spready-work\\outputs\\wc2026-elo\\wc2026.spready";
const DEFAULT_SHEET_NAME = "Matches";
const DEFAULT_EDIT_COLUMN = "Goals A";
const DEFAULT_ITERATIONS = 20;
const DEFAULT_WARMUP_ITERATIONS = 3;
const DEFAULT_VISIBLE_COLUMN_COUNT = 10;
const DEFAULT_VISIBLE_ROW_COUNT = 36;
const VISIBLE_COLUMN_PADDING = 4;
const VISIBLE_ROW_PADDING = 24;

function printUsage() {
  console.log(`Usage: npm run perf:cell-editing -- [options]

Options:
  --workbook PATH       .spready workbook path
                        default: ${DEFAULT_WORKBOOK_PATH}
  --sheet NAME          sheet name to edit, case-insensitive
                        default: ${DEFAULT_SHEET_NAME}
  --edit-column NAME    header name of the edited column
                        default: ${DEFAULT_EDIT_COLUMN}
  --iterations N        measured edit iterations
                        default: ${DEFAULT_ITERATIONS}
  --warmup N            unmeasured warmup edit iterations
                        default: ${DEFAULT_WARMUP_ITERATIONS}
  --view-x N            visible region x
                        default: 0
  --view-y N            visible region y
                        default: 0
  --view-width N        visible region width before padding
                        default: ${DEFAULT_VISIBLE_COLUMN_COUNT}
  --view-height N       visible region height before padding
                        default: ${DEFAULT_VISIBLE_ROW_COUNT}
  --order MODE          selected-first or range-first
                        default: selected-first
  --json                print only machine-readable JSON
  --help                show this help
`);
}

function parseIntegerOption(name, value, minimum = 0) {
  const parsed = Number.parseInt(value ?? "", 10);

  if (!Number.isInteger(parsed) || parsed < minimum) {
    throw new Error(`${name} must be an integer >= ${minimum}.`);
  }

  return parsed;
}

function parseArgs(argv) {
  const options = {
    editColumnName: DEFAULT_EDIT_COLUMN,
    iterations: DEFAULT_ITERATIONS,
    json: false,
    order: "selected-first",
    sheetName: DEFAULT_SHEET_NAME,
    viewHeight: DEFAULT_VISIBLE_ROW_COUNT,
    viewWidth: DEFAULT_VISIBLE_COLUMN_COUNT,
    viewX: 0,
    viewY: 0,
    warmupIterations: DEFAULT_WARMUP_ITERATIONS,
    workbookPath: DEFAULT_WORKBOOK_PATH,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    switch (arg) {
      case "--workbook":
        options.workbookPath = argv[++index];
        break;
      case "--sheet":
        options.sheetName = argv[++index];
        break;
      case "--edit-column":
        options.editColumnName = argv[++index];
        break;
      case "--iterations":
        options.iterations = parseIntegerOption("--iterations", argv[++index], 1);
        break;
      case "--warmup":
        options.warmupIterations = parseIntegerOption("--warmup", argv[++index], 0);
        break;
      case "--view-x":
        options.viewX = parseIntegerOption("--view-x", argv[++index], 0);
        break;
      case "--view-y":
        options.viewY = parseIntegerOption("--view-y", argv[++index], 0);
        break;
      case "--view-width":
        options.viewWidth = parseIntegerOption("--view-width", argv[++index], 1);
        break;
      case "--view-height":
        options.viewHeight = parseIntegerOption("--view-height", argv[++index], 1);
        break;
      case "--order":
        options.order = argv[++index];
        if (!["selected-first", "range-first"].includes(options.order)) {
          throw new Error('--order must be "selected-first" or "range-first".');
        }
        break;
      case "--json":
        options.json = true;
        break;
      case "--help":
      case "-h":
        printUsage();
        process.exit(0);
        break;
      default:
        throw new Error(`Unknown option: ${arg}`);
    }
  }

  if (!options.workbookPath) {
    throw new Error("--workbook is required.");
  }

  return options;
}

function normalizeName(value) {
  return value.trim().toLowerCase();
}

function buildVisibleRangeRequest(sheet, options) {
  const startColumn = Math.max(0, options.viewX - VISIBLE_COLUMN_PADDING);
  const startRow = Math.max(0, options.viewY - VISIBLE_ROW_PADDING);

  return {
    columnCount: Math.max(
      1,
      Math.min(sheet.columnCount - startColumn, options.viewWidth + VISIBLE_COLUMN_PADDING * 2),
    ),
    rowCount: Math.max(
      1,
      Math.min(sheet.rowCount - startRow, options.viewHeight + VISIBLE_ROW_PADDING * 2),
    ),
    sheetId: sheet.id,
    startColumn,
    startRow,
  };
}

function timeMs(callback) {
  const startedMs = performance.now();
  const value = callback();

  return {
    durationMs: performance.now() - startedMs,
    value,
  };
}

async function timeAsyncMs(callback) {
  const startedMs = performance.now();
  const value = await callback();

  return {
    durationMs: performance.now() - startedMs,
    value,
  };
}

function parseTraceLog(message) {
  const fields = {};

  for (const token of message.split(" ")) {
    const separatorIndex = token.indexOf("=");

    if (separatorIndex < 0) {
      continue;
    }

    const key = token.slice(0, separatorIndex);
    const rawValue = token.slice(separatorIndex + 1);
    const numericValue = Number(rawValue);

    fields[key] = Number.isFinite(numericValue) && rawValue.trim() !== "" ? numericValue : rawValue;
  }

  return fields;
}

function getQuantile(sortedValues, quantile) {
  if (sortedValues.length === 0) {
    return 0;
  }

  const index = Math.min(
    sortedValues.length - 1,
    Math.max(0, Math.ceil(sortedValues.length * quantile) - 1),
  );

  return sortedValues[index];
}

function summarize(values) {
  const sorted = [...values].sort((left, right) => left - right);
  const total = values.reduce((sum, value) => sum + value, 0);

  return {
    avgMs: values.length === 0 ? 0 : total / values.length,
    maxMs: sorted[sorted.length - 1] ?? 0,
    minMs: sorted[0] ?? 0,
    p50Ms: getQuantile(sorted, 0.5),
    p95Ms: getQuantile(sorted, 0.95),
    totalMs: total,
  };
}

function roundStats(stats) {
  return Object.fromEntries(
    Object.entries(stats).map(([key, value]) => [
      key,
      typeof value === "number" ? Number(value.toFixed(3)) : value,
    ]),
  );
}

function aggregatePhaseStats(iterations) {
  const phaseNames = [
    "applyTransaction",
    "getCellData",
    "getSheetRange",
    "getSheetDisplayRange",
    "getSheetStyleRange",
    "aftercare",
    "iteration",
  ];

  return Object.fromEntries(
    phaseNames.map((phaseName) => [
      phaseName,
      roundStats(summarize(iterations.map((iteration) => iteration.timings[phaseName]))),
    ]),
  );
}

function aggregateEvaluationStats(iterations) {
  const totals = {
    coldEvaluations: 0,
    cellsEvaluated: 0,
    dependencyKeysRecorded: 0,
    durationMs: 0,
    formulasParsed: 0,
    rangeCellsMaterialized: 0,
  };
  const byReason = {};

  for (const iteration of iterations) {
    for (const evaluation of iteration.evaluations) {
      totals.coldEvaluations += 1;
      totals.cellsEvaluated += Number(evaluation.cellsEvaluated ?? 0);
      totals.dependencyKeysRecorded += Number(evaluation.dependencyKeysRecorded ?? 0);
      totals.durationMs += Number(evaluation.durationMs ?? 0);
      totals.formulasParsed += Number(evaluation.formulasParsed ?? 0);
      totals.rangeCellsMaterialized += Number(evaluation.rangeCellsMaterialized ?? 0);

      const reason = String(evaluation.reason ?? "unknown");
      byReason[reason] ??= {
        coldEvaluations: 0,
        cellsEvaluated: 0,
        dependencyKeysRecorded: 0,
        durationMs: 0,
        formulasParsed: 0,
        rangeCellsMaterialized: 0,
      };
      byReason[reason].coldEvaluations += 1;
      byReason[reason].cellsEvaluated += Number(evaluation.cellsEvaluated ?? 0);
      byReason[reason].dependencyKeysRecorded += Number(evaluation.dependencyKeysRecorded ?? 0);
      byReason[reason].durationMs += Number(evaluation.durationMs ?? 0);
      byReason[reason].formulasParsed += Number(evaluation.formulasParsed ?? 0);
      byReason[reason].rangeCellsMaterialized += Number(evaluation.rangeCellsMaterialized ?? 0);
    }
  }

  return {
    byReason,
    totals: {
      ...totals,
      avgDurationMs:
        totals.coldEvaluations === 0
          ? 0
          : Number((totals.durationMs / totals.coldEvaluations).toFixed(3)),
      durationMs: Number(totals.durationMs.toFixed(3)),
    },
  };
}

function createNextCellValue(currentValue, iterationIndex) {
  const parsed = Number(currentValue);

  if (Number.isFinite(parsed)) {
    return String((Math.trunc(parsed) + 1 + (iterationIndex % 2)) % 10);
  }

  return `${currentValue || "x"}*`;
}

function findSheet(summary, sheetName) {
  const exactName = normalizeName(sheetName);
  const sheet = summary.sheets.find((entry) => normalizeName(entry.name) === exactName);

  if (!sheet) {
    throw new Error(
      `Sheet "${sheetName}" was not found. Available sheets: ${summary.sheets
        .map((entry) => entry.name)
        .join(", ")}`,
    );
  }

  return sheet;
}

function findColumnIndex(headerRow, columnName) {
  const exactName = normalizeName(columnName);
  const columnIndex = headerRow.findIndex((value) => normalizeName(value) === exactName);

  if (columnIndex < 0) {
    throw new Error(`Column "${columnName}" was not found. Header row: ${headerRow.join(", ")}`);
  }

  return columnIndex;
}

function findEditableRows(rawRange, columnIndex) {
  const rows = [];

  for (let rowIndex = 1; rowIndex < rawRange.values.length; rowIndex += 1) {
    const value = rawRange.values[rowIndex]?.[columnIndex] ?? "";

    if (value !== "" && !value.startsWith("=")) {
      rows.push(rowIndex);
    }
  }

  return rows;
}

function runUiEditCycle(controller, sheet, visibleRequest, edit, options) {
  const timings = {};
  const iterationStartedMs = performance.now();

  timings.applyTransaction = timeMs(() =>
    controller.applyTransaction({
      operations: [
        {
          columnIndex: edit.columnIndex,
          rowIndex: edit.rowIndex,
          sheetId: sheet.id,
          type: "setCell",
          value: edit.value,
        },
      ],
    }),
  ).durationMs;

  const runSelectedCell = () => {
    timings.getCellData = timeMs(() =>
      controller.getCellData({
        columnIndex: edit.columnIndex,
        rowIndex: edit.rowIndex,
        sheetId: sheet.id,
      }),
    ).durationMs;
  };

  const runVisibleRange = () => {
    timings.getSheetRange = timeMs(() => controller.getSheetRange(visibleRequest)).durationMs;
    timings.getSheetDisplayRange = timeMs(() =>
      controller.getSheetDisplayRange(visibleRequest),
    ).durationMs;
    timings.getSheetStyleRange = timeMs(() =>
      controller.getSheetStyleRange(visibleRequest),
    ).durationMs;
  };

  const aftercareStartedMs = performance.now();

  if (options.order === "range-first") {
    runVisibleRange();
    runSelectedCell();
  } else {
    runSelectedCell();
    runVisibleRange();
  }

  timings.aftercare = performance.now() - aftercareStartedMs;
  timings.iteration = performance.now() - iterationStartedMs;
  return timings;
}

function printHumanReport(report) {
  console.log("Spready cell editing benchmark");
  console.log(`Workbook: ${report.workbookPath}`);
  console.log(
    `Sheet: ${report.sheet.name} (${report.sheet.rowCount} rows x ${report.sheet.columnCount} cols)`,
  );
  console.log(
    `Edit target: ${report.editColumnName} column=${report.editColumnIndex} rows=${report.editRowsSample.join(", ")}`,
  );
  console.log(
    `Visible request: row=${report.visibleRequest.startRow} col=${report.visibleRequest.startColumn} rows=${report.visibleRequest.rowCount} cols=${report.visibleRequest.columnCount}`,
  );
  console.log(
    `Iterations: ${report.iterations.length} measured, ${report.warmupIterations} warmup, order=${report.order}`,
  );
  console.log("");
  console.table(report.phaseStats);
  console.log("Evaluation totals");
  console.table(report.evaluationStats.byReason);
  console.log(report.evaluationStats.totals);
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const workbookPath = path.resolve(options.workbookPath);

  if (!existsSync(workbookPath)) {
    throw new Error(`Workbook file was not found: ${workbookPath}`);
  }

  const workbookControllerModule = await import("../src/workbook-controller.ts");
  const { WorkbookController } = workbookControllerModule.default ?? workbookControllerModule;
  const traceLogs = [];
  const controller = new WorkbookController({
    evaluationTraceSink: (message) => traceLogs.push(parseTraceLog(message)),
    performanceClock: () => performance.now(),
  });
  const openTiming = await timeAsyncMs(() =>
    controller.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath: workbookPath,
    }),
  );

  const sheet = findSheet(controller.getSummary(), options.sheetName);
  const fullRawRange = controller.getSheetRange({
    columnCount: sheet.columnCount,
    rowCount: sheet.rowCount,
    sheetId: sheet.id,
    startColumn: 0,
    startRow: 0,
  });
  const headerRow = fullRawRange.values[0] ?? [];
  const editColumnIndex = findColumnIndex(headerRow, options.editColumnName);
  const editRows = findEditableRows(fullRawRange, editColumnIndex);

  if (editRows.length === 0) {
    throw new Error(`No editable rows found in column "${options.editColumnName}".`);
  }

  const visibleRequest = buildVisibleRangeRequest(sheet, options);
  controller.getSheetDisplayRange(visibleRequest);
  controller.getCellData({
    columnIndex: editColumnIndex,
    rowIndex: editRows[0],
    sheetId: sheet.id,
  });
  traceLogs.length = 0;

  const totalIterations = options.warmupIterations + options.iterations;
  const measuredIterations = [];

  for (let iterationIndex = 0; iterationIndex < totalIterations; iterationIndex += 1) {
    const rowIndex = editRows[iterationIndex % editRows.length];
    const currentValue =
      controller.getSheetRange({
        columnCount: 1,
        rowCount: 1,
        sheetId: sheet.id,
        startColumn: editColumnIndex,
        startRow: rowIndex,
      }).values[0]?.[0] ?? "";
    const edit = {
      columnIndex: editColumnIndex,
      rowIndex,
      value: createNextCellValue(currentValue, iterationIndex),
    };
    const traceStartIndex = traceLogs.length;
    const timings = runUiEditCycle(controller, sheet, visibleRequest, edit, options);
    const evaluations = traceLogs.slice(traceStartIndex);

    if (iterationIndex >= options.warmupIterations) {
      measuredIterations.push({
        edit,
        evaluations,
        timings: Object.fromEntries(
          Object.entries(timings).map(([key, value]) => [key, Number(value.toFixed(3))]),
        ),
      });
    }
  }

  const report = {
    editColumnIndex,
    editColumnName: options.editColumnName,
    editRowsSample: editRows.slice(0, Math.min(10, editRows.length)),
    evaluationStats: aggregateEvaluationStats(measuredIterations),
    iterations: measuredIterations,
    openWorkbookMs: Number(openTiming.durationMs.toFixed(3)),
    order: options.order,
    phaseStats: aggregatePhaseStats(measuredIterations),
    sheet,
    visibleRequest,
    warmupIterations: options.warmupIterations,
    workbookPath,
  };

  if (options.json) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    printHumanReport(report);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
