import DataEditor, {
  CompactSelection,
  GridCellKind,
  type DataEditorRef,
  type EditableGridCell,
  type GridCell,
  type GridColumn,
  type GridSelection,
  type Item,
  type Theme,
} from "@glideapps/glide-data-grid";
import {
  type ChangeEvent,
  type KeyboardEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { flushSync } from "react-dom";

import { APP_MENU_ACTIONS, type AppMenuAction } from "./app-menu";
import { type ChartEditorWindowRequest } from "./chart-editor-state";
import { ChartEditorDialog } from "./ChartEditorWindow";
import { WorkbookChartOverlay } from "./WorkbookChartOverlay";
import {
  getColumnTitle,
  isFormulaInput,
  parseTsv,
  serializeTsv,
  type CellDataResult,
  type ClipboardRangeMode,
  type WorkbookChartLayout,
  type WorkbookSheetChartPreviewsResult,
  type SheetDisplayRangeResult,
  type SheetRangeRequest,
  type SheetRangeResult,
  type WorkbookSummary,
  type WorkbookTransactionOperation,
} from "./workbook-core";
import { ToastViewport } from "./ToastViewport";
import {
  enqueueToast,
  removeToast,
  type ToastNotification,
} from "./toast-state";

const DEFAULT_COLUMN_WIDTH = 140;
const DEFAULT_VISIBLE_COLUMN_COUNT = 10;
const DEFAULT_VISIBLE_ROW_COUNT = 36;
const DEFAULT_WORKBOOK_FILE_NAME = "Workbook.spready";
const VISIBLE_COLUMN_PADDING = 4;
const VISIBLE_ROW_PADDING = 24;

const GRID_THEME: Partial<Theme> = {
  accentColor: "#2563eb",
  accentFg: "#ffffff",
  accentLight: "rgba(37, 99, 235, 0.16)",
  bgCell: "#ffffff",
  bgCellMedium: "#f8fafc",
  bgHeader: "#f3f6fb",
  bgHeaderHasFocus: "#eaf1ff",
  bgHeaderHovered: "#eef4ff",
  bgBubble: "#eaf1ff",
  bgBubbleSelected: "#2563eb",
  bgIconHeader: "#e2e8f0",
  bgSearchResult: "#dbeafe",
  borderColor: "#cbd5e1",
  drilldownBorder: "#cbd5e1",
  fgIconHeader: "#475569",
  headerBottomBorderColor: "#cbd5e1",
  horizontalBorderColor: "#e2e8f0",
  linkColor: "#2563eb",
  resizeIndicatorColor: "#2563eb",
  roundingRadius: 0,
  textBubble: "#0f172a",
  textDark: "#0f172a",
  textHeader: "#334155",
  textHeaderSelected: "#0f172a",
  textLight: "#94a3b8",
  textMedium: "#475569",
};

type VisibleRegion = {
  height: number;
  width: number;
  x: number;
  y: number;
};

type ChartEditorSession = {
  expectedVersion: number;
  request: ChartEditorWindowRequest;
};

type RangeCache = SheetDisplayRangeResult | SheetRangeResult;

function buildRangeRequest(
  activeSheetId: string,
  columnCount: number,
  rowCount: number,
  region: VisibleRegion | null,
): SheetRangeRequest {
  const targetRegion = region ?? {
    height: Math.min(rowCount, DEFAULT_VISIBLE_ROW_COUNT),
    width: Math.min(columnCount, DEFAULT_VISIBLE_COLUMN_COUNT),
    x: 0,
    y: 0,
  };
  const startColumn = Math.max(0, targetRegion.x - VISIBLE_COLUMN_PADDING);
  const startRow = Math.max(0, targetRegion.y - VISIBLE_ROW_PADDING);

  return {
    columnCount: Math.max(
      1,
      Math.min(
        columnCount - startColumn,
        targetRegion.width + VISIBLE_COLUMN_PADDING * 2,
      ),
    ),
    rowCount: Math.max(
      1,
      Math.min(
        rowCount - startRow,
        targetRegion.height + VISIBLE_ROW_PADDING * 2,
      ),
    ),
    sheetId: activeSheetId,
    startColumn,
    startRow,
  };
}

function createColumns(columnCount: number): GridColumn[] {
  return Array.from({ length: columnCount }, (_, index) => ({
    id: `column-${index}`,
    title: getColumnTitle(index),
    width: DEFAULT_COLUMN_WIDTH,
  }));
}

function createEmptyGridSelection(): GridSelection {
  return {
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  };
}

function createCellSelection(cell: Item): GridSelection {
  const [columnIndex, rowIndex] = cell;

  return {
    columns: CompactSelection.empty(),
    current: {
      cell,
      range: {
        height: 1,
        width: 1,
        x: columnIndex,
        y: rowIndex,
      },
      rangeStack: [],
    },
    rows: CompactSelection.empty(),
  };
}

function createLoadingCell(): GridCell {
  return {
    allowOverlay: false,
    kind: GridCellKind.Loading,
  };
}

function createTextCell(input: string, display: string): GridCell {
  return {
    allowOverlay: true,
    copyData: input,
    data: input,
    displayData: display,
    kind: GridCellKind.Text,
  };
}

function getCachedCellValue(
  cache: RangeCache | null,
  columnIndex: number,
  rowIndex: number,
  sheetId?: string,
): string | undefined {
  if (!cache || cache.sheetId !== sheetId) {
    return undefined;
  }

  if (rowIndex < cache.startRow || columnIndex < cache.startColumn) {
    return undefined;
  }

  const rowOffset = rowIndex - cache.startRow;
  const columnOffset = columnIndex - cache.startColumn;

  if (rowOffset >= cache.rowCount || columnOffset >= cache.columnCount) {
    return undefined;
  }

  return cache.values[rowOffset]?.[columnOffset];
}

function setCachedCellValue<Cache extends RangeCache>(
  cache: Cache | null,
  columnIndex: number,
  rowIndex: number,
  sheetId: string,
  value: string,
): Cache | null {
  if (!cache || cache.sheetId !== sheetId) {
    return cache;
  }

  if (
    rowIndex < cache.startRow ||
    columnIndex < cache.startColumn ||
    rowIndex >= cache.startRow + cache.rowCount ||
    columnIndex >= cache.startColumn + cache.columnCount
  ) {
    return cache;
  }

  const rowOffset = rowIndex - cache.startRow;
  const columnOffset = columnIndex - cache.startColumn;
  const nextValues = [...cache.values];
  const nextRow = [...(nextValues[rowOffset] ?? [])];

  nextRow[columnOffset] = value;
  nextValues[rowOffset] = nextRow;

  return {
    ...cache,
    values: nextValues,
  };
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : "Unknown error";
}

function getPastedCellValue(
  target: Item,
  selectedCell: Item | null,
  values: readonly (readonly string[])[],
): string | null {
  if (!selectedCell) {
    return null;
  }

  const [startColumn, startRow] = target;
  const [selectedColumn, selectedRow] = selectedCell;
  const rowOffset = selectedRow - startRow;
  const columnOffset = selectedColumn - startColumn;

  if (
    rowOffset < 0 ||
    columnOffset < 0 ||
    rowOffset >= values.length ||
    columnOffset >= values[rowOffset].length
  ) {
    return null;
  }

  return values[rowOffset][columnOffset] ?? "";
}

function getSelectedCellAddress(selectedCell: Item | null): string {
  if (!selectedCell) {
    return "";
  }

  return `${getColumnTitle(selectedCell[0])}${selectedCell[1] + 1}`;
}

function getCurrentSelectionRange(
  selection: GridSelection,
  sheetId: string,
): SheetRangeRequest | null {
  const range = selection.current?.range;

  if (!range) {
    return null;
  }

  return {
    columnCount: Math.max(1, range.width),
    rowCount: Math.max(1, range.height),
    sheetId,
    startColumn: range.x,
    startRow: range.y,
  };
}

function selectionContainsCell(selection: GridSelection, cell: Item): boolean {
  const [columnIndex, rowIndex] = cell;
  const range = selection.current?.range;

  if (
    range &&
    columnIndex >= range.x &&
    columnIndex < range.x + range.width &&
    rowIndex >= range.y &&
    rowIndex < range.y + range.height
  ) {
    return true;
  }

  return false;
}

function compactSelectionToRuns(
  selection: CompactSelection,
): Array<{ count: number; start: number }> {
  const runs: Array<{ count: number; start: number }> = [];
  let runStart: number | null = null;
  let previousIndex: number | null = null;

  for (const index of selection) {
    if (runStart === null) {
      runStart = index;
      previousIndex = index;
      continue;
    }

    if (previousIndex !== null && index === previousIndex + 1) {
      previousIndex = index;
      continue;
    }

    runs.push({
      count: (previousIndex ?? runStart) - runStart + 1,
      start: runStart,
    });
    runStart = index;
    previousIndex = index;
  }

  if (runStart !== null) {
    runs.push({
      count: (previousIndex ?? runStart) - runStart + 1,
      start: runStart,
    });
  }

  return runs;
}

function getClearSelectionOperations(
  selection: GridSelection,
  sheetId: string,
  rowCount: number,
  columnCount: number,
): WorkbookTransactionOperation[] {
  const operations: WorkbookTransactionOperation[] = [];
  const range = selection.current?.range;

  if (range) {
    operations.push({
      columnCount: Math.max(1, range.width),
      rowCount: Math.max(1, range.height),
      sheetId,
      startColumn: range.x,
      startRow: range.y,
      type: "clearRange",
    });
  }

  for (const rowRun of compactSelectionToRuns(selection.rows)) {
    operations.push({
      columnCount,
      rowCount: rowRun.count,
      sheetId,
      startColumn: 0,
      startRow: rowRun.start,
      type: "clearRange",
    });
  }

  for (const columnRun of compactSelectionToRuns(selection.columns)) {
    operations.push({
      columnCount: columnRun.count,
      rowCount,
      sheetId,
      startColumn: columnRun.start,
      startRow: 0,
      type: "clearRange",
    });
  }

  return operations;
}

function cellWouldBeCleared(
  selection: GridSelection,
  cell: Item | null,
): boolean {
  if (!cell) {
    return false;
  }

  if (selectionContainsCell(selection, cell)) {
    return true;
  }

  return (
    selection.columns.hasIndex(cell[0]) || selection.rows.hasIndex(cell[1])
  );
}

function replaceInputSelection(
  input: HTMLInputElement,
  nextText: string,
): { selectionStart: number; value: string } {
  const selectionStart = input.selectionStart ?? input.value.length;
  const selectionEnd = input.selectionEnd ?? input.value.length;
  const value =
    input.value.slice(0, selectionStart) +
    nextText +
    input.value.slice(selectionEnd);

  return {
    selectionStart: selectionStart + nextText.length,
    value,
  };
}

function isEditableShortcutTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) {
    return false;
  }

  if (target.isContentEditable) {
    return true;
  }

  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLSelectElement ||
    target instanceof HTMLTextAreaElement
  );
}

function getDefaultWorkbookFilePath(summary: WorkbookSummary | null): string {
  return summary?.documentFilePath ?? DEFAULT_WORKBOOK_FILE_NAME;
}

export default function App() {
  const [chartEditorSession, setChartEditorSession] =
    useState<ChartEditorSession | null>(null);
  const [formulaInputValue, setFormulaInputValue] = useState("");
  const [gridSelection, setGridSelection] = useState<GridSelection>(
    createEmptyGridSelection,
  );
  const [gridViewportNonce, setGridViewportNonce] = useState(0);
  const [isSheetChartPreviewsLoading, setIsSheetChartPreviewsLoading] =
    useState(false);
  const [selectedChartId, setSelectedChartId] = useState<string | null>(null);
  const [selectedCellData, setSelectedCellData] =
    useState<CellDataResult | null>(null);
  const [sheetChartPreviews, setSheetChartPreviews] =
    useState<WorkbookSheetChartPreviewsResult | null>(null);
  const [sheetSummary, setSheetSummary] = useState<WorkbookSummary | null>(
    null,
  );
  const [toasts, setToasts] = useState<ToastNotification[]>([]);
  const [viewNonce, setViewNonce] = useState(0);

  const exportPathRef = useRef<string>();
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const displayRangeCacheRef = useRef<SheetDisplayRangeResult | null>(null);
  const gridRef = useRef<DataEditorRef>(null);
  const lastVisibleRegionRef = useRef<VisibleRegion | null>(null);
  const pendingCellDataRequestIdRef = useRef(0);
  const pendingRangeRequestIdRef = useRef(0);
  const pendingSheetChartPreviewsRequestIdRef = useRef(0);
  const rawRangeCacheRef = useRef<SheetRangeResult | null>(null);
  const sheetSurfaceRef = useRef<HTMLElement>(null);
  const isChartEditorOpen = chartEditorSession !== null;

  const activeSheet = useMemo(
    () =>
      sheetSummary?.sheets.find(
        (sheet) => sheet.id === sheetSummary.activeSheetId,
      ) ?? null,
    [sheetSummary],
  );
  const selectedCell = gridSelection.current?.cell ?? null;
  const selectedCellAddress = useMemo(
    () => getSelectedCellAddress(selectedCell),
    [selectedCell],
  );
  const activeSheetChartEntries = useMemo(() => {
    const activeSheetCharts = (sheetSummary?.charts ?? []).filter(
      (chart) => chart.sheetId === activeSheet?.id,
    );

    if (!sheetChartPreviews) {
      return activeSheetCharts.map((chart) => ({
        chartType: chart.chartType,
        id: chart.id,
        layout: chart.layout,
        name: chart.name,
        status: chart.status,
      }));
    }

    const chartStatuses = new Map(
      activeSheetCharts.map((chart) => [chart.id, chart.status]),
    );

    return sheetChartPreviews.previews.map((preview) => ({
      chartType: preview.chart.spec.chartType,
      id: preview.chart.id,
      layout: preview.chart.layout,
      name: preview.chart.name,
      status: chartStatuses.get(preview.chart.id) ?? preview.status,
    }));
  }, [activeSheet?.id, sheetChartPreviews, sheetSummary?.charts]);
  const rowCount = activeSheet?.rowCount ?? 1;
  const columnCount = activeSheet?.columnCount ?? 1;
  const columns = useMemo(() => createColumns(columnCount), [columnCount]);
  const currentSelectionRange = useMemo(
    () =>
      activeSheet
        ? getCurrentSelectionRange(gridSelection, activeSheet.id)
        : null,
    [activeSheet, gridSelection],
  );
  const dismissToast = useCallback((toastId: string) => {
    setToasts((current) => removeToast(current, toastId));
  }, []);
  const pushErrorToast = useCallback((error: unknown) => {
    setToasts((current) =>
      enqueueToast(current, {
        kind: "error",
        title: getErrorMessage(error),
      }),
    );
  }, []);

  const applyTransaction = useCallback(
    async (
      operations: Parameters<
        typeof window.appShell.applyTransaction
      >[0]["operations"],
    ) => {
      const result = await window.appShell.applyTransaction({ operations });

      setSheetSummary(result.summary);

      return result;
    },
    [],
  );

  const loadVisibleRange = useCallback(
    async (region: VisibleRegion | null) => {
      if (!activeSheet) {
        return;
      }

      const request = buildRangeRequest(
        activeSheet.id,
        activeSheet.columnCount,
        activeSheet.rowCount,
        region,
      );
      const requestId = pendingRangeRequestIdRef.current + 1;

      pendingRangeRequestIdRef.current = requestId;

      try {
        const [rawRange, displayRange] = await Promise.all([
          window.appShell.getSheetRange(request),
          window.appShell.getSheetDisplayRange(request),
        ]);

        if (pendingRangeRequestIdRef.current !== requestId) {
          return;
        }

        rawRangeCacheRef.current = rawRange;
        displayRangeCacheRef.current = displayRange;
        setViewNonce((current) => current + 1);
      } catch (error) {
        pushErrorToast(error);
      }
    },
    [activeSheet, pushErrorToast],
  );

  const refreshSelectedCellData = useCallback(async () => {
    if (!activeSheet || !selectedCell) {
      setSelectedCellData(null);
      setFormulaInputValue("");
      return;
    }

    const requestId = pendingCellDataRequestIdRef.current + 1;

    pendingCellDataRequestIdRef.current = requestId;

    try {
      const [columnIndex, rowIndex] = selectedCell;
      const cellData = await window.appShell.getCellData({
        columnIndex,
        rowIndex,
        sheetId: activeSheet.id,
      });

      if (pendingCellDataRequestIdRef.current !== requestId) {
        return;
      }

      setSelectedCellData(cellData);
      setFormulaInputValue(cellData.input);
    } catch (error) {
      pushErrorToast(error);
    }
  }, [activeSheet, pushErrorToast, selectedCell]);

  const getCellContent = useCallback(
    (cell: Item): GridCell => {
      const [columnIndex, rowIndex] = cell;
      const rawValue = getCachedCellValue(
        rawRangeCacheRef.current,
        columnIndex,
        rowIndex,
        activeSheet?.id,
      );
      const displayValue = getCachedCellValue(
        displayRangeCacheRef.current,
        columnIndex,
        rowIndex,
        activeSheet?.id,
      );

      if (rawValue === undefined || displayValue === undefined) {
        return createLoadingCell();
      }

      return createTextCell(rawValue, displayValue);
    },
    [activeSheet?.id, viewNonce],
  );

  const getCellsForSelection = useCallback(
    (selection: VisibleRegion) => {
      return async () => {
        if (!activeSheet) {
          return [];
        }

        const request = {
          columnCount: selection.width,
          rowCount: selection.height,
          sheetId: activeSheet.id,
          startColumn: selection.x,
          startRow: selection.y,
        };
        const [rawRange, displayRange] = await Promise.all([
          window.appShell.getSheetRange(request),
          window.appShell.getSheetDisplayRange(request),
        ]);

        return displayRange.values.map((row, rowOffset) =>
          row.map((displayValue, columnOffset) =>
            createTextCell(
              rawRange.values[rowOffset]?.[columnOffset] ?? displayValue,
              displayValue,
            ),
          ),
        );
      };
    },
    [activeSheet],
  );

  const handleCellEdited = useCallback(
    (cell: Item, newValue: EditableGridCell) => {
      if (newValue.kind !== GridCellKind.Text || !activeSheet) {
        return;
      }

      const [columnIndex, rowIndex] = cell;

      rawRangeCacheRef.current = setCachedCellValue(
        rawRangeCacheRef.current,
        columnIndex,
        rowIndex,
        activeSheet.id,
        newValue.data,
      );
      displayRangeCacheRef.current = setCachedCellValue(
        displayRangeCacheRef.current,
        columnIndex,
        rowIndex,
        activeSheet.id,
        newValue.data,
      );

      if (selectedCell?.[0] === columnIndex && selectedCell?.[1] === rowIndex) {
        setFormulaInputValue(newValue.data);
        setSelectedCellData((current) =>
          current
            ? {
                ...current,
                display: newValue.data,
                errorCode: undefined,
                input: newValue.data,
                isFormula: isFormulaInput(newValue.data),
              }
            : current,
        );
      }

      setViewNonce((current) => current + 1);

      void applyTransaction([
        {
          columnIndex,
          rowIndex,
          type: "setCell",
          value: newValue.data,
        },
      ]).catch((error) => {
        pushErrorToast(error);
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
      });
    },
    [
      activeSheet,
      applyTransaction,
      loadVisibleRange,
      pushErrorToast,
      refreshSelectedCellData,
      selectedCell,
    ],
  );

  const handlePaste = useCallback(
    (target: Item, values: readonly (readonly string[])[]) => {
      if (!activeSheet || values.length === 0) {
        return false;
      }

      const [startColumn, startRow] = target;
      const nextValues = values.map((row) => [...row]);

      for (let rowOffset = 0; rowOffset < nextValues.length; rowOffset += 1) {
        for (
          let columnOffset = 0;
          columnOffset < nextValues[rowOffset].length;
          columnOffset += 1
        ) {
          const nextValue = nextValues[rowOffset][columnOffset] ?? "";

          rawRangeCacheRef.current = setCachedCellValue(
            rawRangeCacheRef.current,
            startColumn + columnOffset,
            startRow + rowOffset,
            activeSheet.id,
            nextValue,
          );
          displayRangeCacheRef.current = setCachedCellValue(
            displayRangeCacheRef.current,
            startColumn + columnOffset,
            startRow + rowOffset,
            activeSheet.id,
            nextValue,
          );
        }
      }

      const selectedPastedValue = getPastedCellValue(
        target,
        selectedCell,
        values,
      );

      if (selectedPastedValue !== null) {
        setFormulaInputValue(selectedPastedValue);
        setSelectedCellData((current) =>
          current
            ? {
                ...current,
                display: selectedPastedValue,
                errorCode: undefined,
                input: selectedPastedValue,
                isFormula: isFormulaInput(selectedPastedValue),
              }
            : current,
        );
      }

      setViewNonce((current) => current + 1);

      void applyTransaction([
        {
          startColumn,
          startRow,
          type: "setRange",
          values: nextValues,
        },
      ]).catch((error) => {
        pushErrorToast(error);
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
      });

      return true;
    },
    [
      activeSheet,
      applyTransaction,
      loadVisibleRange,
      pushErrorToast,
      refreshSelectedCellData,
      selectedCell,
    ],
  );

  const replaceFormulaInputSelection = useCallback((nextText: string) => {
    const input = formulaInputRef.current;

    if (!input) {
      return false;
    }

    const nextState = replaceInputSelection(input, nextText);

    setFormulaInputValue(nextState.value);

    requestAnimationFrame(() => {
      input.focus();
      input.setSelectionRange(
        nextState.selectionStart,
        nextState.selectionStart,
      );
    });

    return true;
  }, []);

  const deleteFormulaInputSelection = useCallback(() => {
    const input = formulaInputRef.current;

    if (!input) {
      return false;
    }

    const selectionStart = input.selectionStart ?? input.value.length;
    const selectionEnd = input.selectionEnd ?? input.value.length;

    if (
      selectionStart === selectionEnd &&
      selectionStart >= input.value.length
    ) {
      return false;
    }

    const deleteEnd =
      selectionStart === selectionEnd ? selectionStart + 1 : selectionEnd;
    const value =
      input.value.slice(0, selectionStart) + input.value.slice(deleteEnd);

    setFormulaInputValue(value);

    requestAnimationFrame(() => {
      input.focus();
      input.setSelectionRange(selectionStart, selectionStart);
    });

    return true;
  }, []);

  const copySelection = useCallback(
    async (mode: ClipboardRangeMode) => {
      const input = formulaInputRef.current;

      if (document.activeElement === input && input) {
        const selectionStart = input.selectionStart ?? input.value.length;
        const selectionEnd = input.selectionEnd ?? input.value.length;

        if (selectionStart === selectionEnd) {
          return false;
        }

        await window.appShell.writeClipboard({
          text: input.value.slice(selectionStart, selectionEnd),
        });
        return true;
      }

      if (!currentSelectionRange) {
        return false;
      }

      try {
        const [rawRange, displayRange] = await Promise.all([
          window.appShell.getSheetRange(currentSelectionRange),
          window.appShell.getSheetDisplayRange(currentSelectionRange),
        ]);
        const rawText = serializeTsv(rawRange.values);
        const displayText = serializeTsv(displayRange.values);

        await window.appShell.writeClipboard({
          payload: {
            displayText,
            displayValues: displayRange.values.map((row) => [...row]),
            rawText,
            rawValues: rawRange.values.map((row) => [...row]),
          },
          text: mode === "display" ? displayText : rawText,
        });
        return true;
      } catch (error) {
        pushErrorToast(error);
        return false;
      }
    },
    [currentSelectionRange, pushErrorToast],
  );

  const cutSelection = useCallback(
    async (mode: ClipboardRangeMode) => {
      const input = formulaInputRef.current;

      if (document.activeElement === input && input) {
        const selectionStart = input.selectionStart ?? input.value.length;
        const selectionEnd = input.selectionEnd ?? input.value.length;

        if (selectionStart === selectionEnd) {
          return false;
        }

        await window.appShell.writeClipboard({
          text: input.value.slice(selectionStart, selectionEnd),
        });

        return replaceFormulaInputSelection("");
      }

      if (!currentSelectionRange) {
        return false;
      }

      try {
        const [rawRange, displayRange] = await Promise.all([
          window.appShell.getSheetRange(currentSelectionRange),
          window.appShell.getSheetDisplayRange(currentSelectionRange),
        ]);
        const rawText = serializeTsv(rawRange.values);
        const displayText = serializeTsv(displayRange.values);

        await window.appShell.writeClipboard({
          payload: {
            displayText,
            displayValues: displayRange.values.map((row) => [...row]),
            rawText,
            rawValues: rawRange.values.map((row) => [...row]),
          },
          text: mode === "display" ? displayText : rawText,
        });

        const result = await window.appShell.cutRange({
          ...currentSelectionRange,
          mode,
        });

        setSheetSummary(result.summary);

        if (
          selectedCell &&
          selectionContainsCell(gridSelection, selectedCell)
        ) {
          setFormulaInputValue("");
          setSelectedCellData((current) =>
            current
              ? {
                  ...current,
                  display: "",
                  errorCode: undefined,
                  input: "",
                  isFormula: false,
                }
              : current,
          );
        }

        return true;
      } catch (error) {
        pushErrorToast(error);
        return false;
      }
    },
    [
      currentSelectionRange,
      gridSelection,
      pushErrorToast,
      replaceFormulaInputSelection,
      selectedCell,
    ],
  );

  const pasteSelection = useCallback(
    async (mode: ClipboardRangeMode) => {
      const clipboard = await window.appShell.readClipboard();
      const input = formulaInputRef.current;

      if (document.activeElement === input && input) {
        const nextText =
          mode === "display"
            ? (clipboard.payload?.displayText ?? clipboard.text)
            : (clipboard.payload?.rawText ?? clipboard.text);

        if (
          nextText.length === 0 &&
          !clipboard.payload &&
          clipboard.text.length === 0
        ) {
          return false;
        }

        return replaceFormulaInputSelection(nextText);
      }

      if (!selectedCell) {
        return false;
      }

      const values =
        mode === "display"
          ? (clipboard.payload?.displayValues ?? parseTsv(clipboard.text))
          : (clipboard.payload?.rawValues ?? parseTsv(clipboard.text));

      return handlePaste(selectedCell, values);
    },
    [handlePaste, replaceFormulaInputSelection, selectedCell],
  );

  const deleteSelection = useCallback(
    (selection: GridSelection = gridSelection) => {
      const input = formulaInputRef.current;

      if (document.activeElement === input && input) {
        return deleteFormulaInputSelection();
      }

      if (selectedChartId) {
        const chartId = selectedChartId;

        setSelectedChartId(null);

        void applyTransaction([
          {
            chartId,
            type: "deleteChart",
          },
        ]).catch((error) => {
          setSelectedChartId(chartId);
          pushErrorToast(error);
        });

        return true;
      }

      if (!activeSheet) {
        return false;
      }

      const operations = getClearSelectionOperations(
        selection,
        activeSheet.id,
        activeSheet.rowCount,
        activeSheet.columnCount,
      );

      if (operations.length === 0) {
        return false;
      }

      if (cellWouldBeCleared(selection, selectedCell)) {
        setFormulaInputValue("");
        setSelectedCellData((current) =>
          current
            ? {
                ...current,
                display: "",
                errorCode: undefined,
                input: "",
                isFormula: false,
              }
            : current,
        );
      }

      void applyTransaction(operations).catch((error) => {
        pushErrorToast(error);
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
      });

      return true;
    },
    [
      activeSheet,
      applyTransaction,
      deleteFormulaInputSelection,
      gridSelection,
      loadVisibleRange,
      pushErrorToast,
      refreshSelectedCellData,
      selectedCell,
      selectedChartId,
    ],
  );

  const handleCellContextMenu = useCallback(
    (cell: Item, event: { preventDefault?: () => void }) => {
      event.preventDefault?.();

      if (!selectionContainsCell(gridSelection, cell)) {
        flushSync(() => {
          setGridSelection(createCellSelection(cell));
        });
      }

      void window.appShell
        .showCellContextMenu({
          canCopy: true,
          canCut: true,
          canDelete: true,
        })
        .catch((error) => {
          pushErrorToast(error);
        });
    },
    [gridSelection, pushErrorToast],
  );

  const commitFormulaBar = useCallback(async () => {
    if (!activeSheet || !selectedCell) {
      return;
    }

    const [columnIndex, rowIndex] = selectedCell;

    if (formulaInputValue === (selectedCellData?.input ?? "")) {
      return;
    }

    rawRangeCacheRef.current = setCachedCellValue(
      rawRangeCacheRef.current,
      columnIndex,
      rowIndex,
      activeSheet.id,
      formulaInputValue,
    );
    displayRangeCacheRef.current = setCachedCellValue(
      displayRangeCacheRef.current,
      columnIndex,
      rowIndex,
      activeSheet.id,
      formulaInputValue,
    );
    setSelectedCellData((current) =>
      current
        ? {
            ...current,
            display: formulaInputValue,
            errorCode: undefined,
            input: formulaInputValue,
            isFormula: isFormulaInput(formulaInputValue),
          }
        : current,
    );
    setViewNonce((current) => current + 1);

    try {
      await applyTransaction([
        {
          columnIndex,
          rowIndex,
          type: "setCell",
          value: formulaInputValue,
        },
      ]);
    } catch (error) {
      pushErrorToast(error);
      void loadVisibleRange(lastVisibleRegionRef.current);
      void refreshSelectedCellData();
    }
  }, [
    activeSheet,
    applyTransaction,
    formulaInputValue,
    loadVisibleRange,
    pushErrorToast,
    refreshSelectedCellData,
    selectedCell,
    selectedCellData?.input,
  ]);

  const addColumn = useCallback(() => {
    if (!activeSheet) {
      return;
    }

    void applyTransaction([
      {
        columnIndex: activeSheet.columnCount,
        count: 1,
        type: "insertColumns",
      },
    ]).catch((error) => {
      pushErrorToast(error);
    });
  }, [activeSheet, applyTransaction, pushErrorToast]);

  const addRow = useCallback(() => {
    if (!activeSheet) {
      return;
    }

    void applyTransaction([
      {
        count: 1,
        rowIndex: activeSheet.rowCount,
        type: "insertRows",
      },
    ]).catch((error) => {
      pushErrorToast(error);
    });
  }, [activeSheet, applyTransaction, pushErrorToast]);

  const addSheet = useCallback(() => {
    void applyTransaction([
      {
        activate: true,
        type: "addSheet",
      },
    ]).catch((error) => {
      pushErrorToast(error);
    });
  }, [applyTransaction, pushErrorToast]);

  const deleteSheet = useCallback(() => {
    if (!activeSheet) {
      return;
    }

    void applyTransaction([
      {
        sheetId: activeSheet.id,
        type: "deleteSheet",
      },
    ]).catch((error) => {
      pushErrorToast(error);
    });
  }, [activeSheet, applyTransaction, pushErrorToast]);

  const handleActiveSheetChange = useCallback(
    (event: ChangeEvent<HTMLSelectElement>) => {
      void applyTransaction([
        {
          sheetId: event.target.value,
          type: "setActiveSheet",
        },
      ]).catch((error) => {
        pushErrorToast(error);
      });
    },
    [applyTransaction, pushErrorToast],
  );

  const handleFormulaInputChange = useCallback(
    (event: ChangeEvent<HTMLInputElement>) => {
      setFormulaInputValue(event.target.value);
    },
    [],
  );

  const handleFormulaInputKeyDown = useCallback(
    (event: KeyboardEvent<HTMLInputElement>) => {
      if (event.key === "Enter") {
        event.preventDefault();
        void commitFormulaBar();
        return;
      }

      if (event.key === "Escape") {
        event.preventDefault();
        setFormulaInputValue(selectedCellData?.input ?? "");
      }
    },
    [commitFormulaBar, selectedCellData?.input],
  );

  const handleImport = useCallback(async () => {
    try {
      const result = await window.appShell.openCsvFile();

      if (result.canceled) {
        return;
      }

      exportPathRef.current = result.filePath;

      await applyTransaction([
        {
          content: result.content,
          sourceFilePath: result.filePath,
          type: "replaceSheetFromCsv",
        },
      ]);
    } catch (error) {
      pushErrorToast(error);
    }
  }, [applyTransaction, pushErrorToast]);

  const handleExport = useCallback(async () => {
    if (!activeSheet) {
      return;
    }

    try {
      const csv = await window.appShell.getSheetCsv(activeSheet.id);
      const defaultPath =
        exportPathRef.current ??
        activeSheet.sourceFilePath ??
        `${activeSheet.name.replaceAll(/\s+/g, "-") || "Sheet"}.csv`;
      const result = await window.appShell.saveCsvFile(csv, defaultPath);

      if (result.canceled) {
        return;
      }

      exportPathRef.current = result.filePath;

      await applyTransaction([
        {
          sourceFilePath: result.filePath,
          type: "setSheetSourceFile",
        },
      ]);
    } catch (error) {
      pushErrorToast(error);
    }
  }, [activeSheet, applyTransaction, pushErrorToast]);

  const handleOpenWorkbook = useCallback(async () => {
    try {
      const result = await window.appShell.openWorkbookFile();

      if (result.canceled) {
        return;
      }

      setSheetSummary(result.summary);
    } catch (error) {
      pushErrorToast(error);
    }
  }, [pushErrorToast]);

  const handleSaveWorkbookAs = useCallback(async () => {
    try {
      const result = await window.appShell.saveWorkbookFileAs(
        getDefaultWorkbookFilePath(sheetSummary),
      );

      if (result.canceled) {
        return;
      }

      setSheetSummary(result.summary);
    } catch (error) {
      pushErrorToast(error);
    }
  }, [pushErrorToast, sheetSummary]);

  const handleSaveWorkbook = useCallback(async () => {
    try {
      if (!sheetSummary?.documentFilePath) {
        await handleSaveWorkbookAs();
        return;
      }

      const result = await window.appShell.saveWorkbookFile(
        sheetSummary.documentFilePath,
      );

      setSheetSummary(result.summary);
    } catch (error) {
      pushErrorToast(error);
    }
  }, [handleSaveWorkbookAs, pushErrorToast, sheetSummary]);

  const closeChartEditor = useCallback(() => {
    setChartEditorSession(null);
  }, []);

  const handleChartEditorVersionConflict = useCallback(
    (message: string) => {
      setChartEditorSession(null);
      pushErrorToast(new Error(message));
    },
    [pushErrorToast],
  );

  const openCreateChartEditor = useCallback(() => {
    if (!activeSheet || !sheetSummary || isChartEditorOpen) {
      return;
    }

    setChartEditorSession({
      expectedVersion: sheetSummary.version,
      request: {
        mode: "create",
        sheetId: activeSheet.id,
      },
    });
  }, [activeSheet, isChartEditorOpen, sheetSummary]);

  const openEditChartEditor = useCallback(
    (chartId: string) => {
      if (!sheetSummary || isChartEditorOpen) {
        return;
      }

      setChartEditorSession({
        expectedVersion: sheetSummary.version,
        request: {
          chartId,
          mode: "edit",
        },
      });
    },
    [isChartEditorOpen, sheetSummary],
  );

  const commitChartLayout = useCallback(
    async (chartId: string, layout: WorkbookChartLayout) => {
      if (!sheetSummary) {
        return;
      }

      const currentMaxZIndex = Math.max(
        -1,
        ...(sheetChartPreviews?.previews ?? [])
          .filter((preview) => preview.chart.id !== chartId)
          .map((preview) => preview.chart.layout.zIndex),
      );

      try {
        const result = await window.appShell.applyTransaction({
          expectedVersion: sheetSummary.version,
          operations: [
            {
              chartId,
              layout: {
                ...layout,
                zIndex: Math.max(layout.zIndex, currentMaxZIndex + 1),
              },
              type: "setChartLayout",
            },
          ],
        });

        setSheetSummary(result.summary);
      } catch (error) {
        pushErrorToast(error);
      }
    },
    [pushErrorToast, sheetChartPreviews?.previews, sheetSummary],
  );

  useEffect(() => {
    let isMounted = true;

    void window.appShell
      .getWorkbookSummary()
      .then((summary) => {
        if (!isMounted) {
          return;
        }

        setSheetSummary(summary);
      })
      .catch((error) => {
        if (!isMounted) {
          return;
        }

        pushErrorToast(error);
      });

    const unsubscribeWorkbook = window.appShell.onWorkbookChanged((summary) => {
      setSheetSummary(summary);
    });

    return () => {
      isMounted = false;
      unsubscribeWorkbook();
    };
  }, [pushErrorToast]);

  useEffect(() => {
    let isCancelled = false;

    void window.appShell
      .setChartDialogOpen(isChartEditorOpen)
      .catch((error) => {
        if (!isCancelled) {
          pushErrorToast(error);
        }
      });

    return () => {
      isCancelled = true;

      if (isChartEditorOpen) {
        void window.appShell.setChartDialogOpen(false);
      }
    };
  }, [isChartEditorOpen, pushErrorToast]);

  useEffect(() => {
    setGridSelection(createEmptyGridSelection());
    setSelectedCellData(null);
    setFormulaInputValue("");
  }, [activeSheet?.id]);

  useEffect(() => {
    if (!selectedCell || !activeSheet) {
      setSelectedCellData(null);
      setFormulaInputValue("");
      return;
    }

    void refreshSelectedCellData();
  }, [
    activeSheet,
    refreshSelectedCellData,
    selectedCell,
    sheetSummary?.version,
  ]);

  useEffect(() => {
    if (!activeSheet) {
      setSheetChartPreviews(null);
      setSelectedChartId(null);
      setIsSheetChartPreviewsLoading(false);
      return;
    }

    rawRangeCacheRef.current = null;
    displayRangeCacheRef.current = null;
    setViewNonce((current) => current + 1);

    exportPathRef.current = activeSheet.sourceFilePath;

    void loadVisibleRange(lastVisibleRegionRef.current);
  }, [
    activeSheet?.columnCount,
    activeSheet?.id,
    activeSheet?.rowCount,
    activeSheet?.sourceFilePath,
    loadVisibleRange,
    sheetSummary?.version,
  ]);

  useEffect(() => {
    if (!activeSheet) {
      return;
    }

    const requestId = pendingSheetChartPreviewsRequestIdRef.current + 1;

    pendingSheetChartPreviewsRequestIdRef.current = requestId;
    setIsSheetChartPreviewsLoading(true);

    void window.appShell
      .getSheetChartPreviews(activeSheet.id)
      .then((result) => {
        if (pendingSheetChartPreviewsRequestIdRef.current !== requestId) {
          return;
        }

        setSheetChartPreviews(result);
        setSelectedChartId((current) =>
          result.previews.some((preview) => preview.chart.id === current)
            ? current
            : null,
        );
      })
      .catch((error) => {
        if (pendingSheetChartPreviewsRequestIdRef.current !== requestId) {
          return;
        }

        setSheetChartPreviews(null);
        setSelectedChartId(null);
        pushErrorToast(error);
      })
      .finally(() => {
        if (pendingSheetChartPreviewsRequestIdRef.current === requestId) {
          setIsSheetChartPreviewsLoading(false);
        }
      });
  }, [activeSheet, pushErrorToast, sheetSummary?.version]);

  useEffect(() => {
    return window.appShell.onMenuAction((action) => {
      if (isChartEditorOpen) {
        return;
      }

      const handleMenuAction = (nextAction: AppMenuAction) => {
        switch (nextAction) {
          case APP_MENU_ACTIONS.cut:
            void cutSelection("raw");
            return;
          case APP_MENU_ACTIONS.cutValues:
            void cutSelection("display");
            return;
          case APP_MENU_ACTIONS.copy:
            void copySelection("raw");
            return;
          case APP_MENU_ACTIONS.copyValues:
            void copySelection("display");
            return;
          case APP_MENU_ACTIONS.openWorkbook:
            void handleOpenWorkbook();
            return;
          case APP_MENU_ACTIONS.saveWorkbook:
            void handleSaveWorkbook();
            return;
          case APP_MENU_ACTIONS.saveWorkbookAs:
            void handleSaveWorkbookAs();
            return;
          case APP_MENU_ACTIONS.importCsv:
            void handleImport();
            return;
          case APP_MENU_ACTIONS.exportCsv:
            void handleExport();
            return;
          case APP_MENU_ACTIONS.paste:
            void pasteSelection("raw");
            return;
          case APP_MENU_ACTIONS.pasteValues:
            void pasteSelection("display");
            return;
          case APP_MENU_ACTIONS.deleteSelection:
            deleteSelection();
            return;
          case APP_MENU_ACTIONS.addRow:
            addRow();
            return;
          case APP_MENU_ACTIONS.addColumn:
            addColumn();
            return;
          case APP_MENU_ACTIONS.newSheet:
            addSheet();
            return;
          case APP_MENU_ACTIONS.deleteSheet:
            deleteSheet();
            return;
          case APP_MENU_ACTIONS.insertChart:
            openCreateChartEditor();
        }
      };

      handleMenuAction(action);
    });
  }, [
    addColumn,
    addRow,
    addSheet,
    cutSelection,
    copySelection,
    deleteSelection,
    deleteSheet,
    handleExport,
    handleImport,
    handleOpenWorkbook,
    handleSaveWorkbook,
    handleSaveWorkbookAs,
    isChartEditorOpen,
    openCreateChartEditor,
    pasteSelection,
  ]);

  useEffect(() => {
    const handleWindowKeyDown = (event: globalThis.KeyboardEvent) => {
      if (isChartEditorOpen) {
        return;
      }

      const isPrimaryModifier = event.ctrlKey || event.metaKey;
      const activeElement = document.activeElement;
      const isFormulaInputFocused = activeElement === formulaInputRef.current;

      if (
        isEditableShortcutTarget(event.target) &&
        !isFormulaInputFocused &&
        event.key !== "Delete"
      ) {
        return;
      }

      if (!isPrimaryModifier && event.key !== "Delete") {
        return;
      }

      if (event.altKey) {
        return;
      }

      const normalizedKey = event.key.toLowerCase();

      if (isPrimaryModifier && normalizedKey === "c") {
        event.preventDefault();
        void copySelection(event.shiftKey ? "display" : "raw");
        return;
      }

      if (isPrimaryModifier && normalizedKey === "x") {
        event.preventDefault();
        void cutSelection(event.shiftKey ? "display" : "raw");
        return;
      }

      if (isPrimaryModifier && normalizedKey === "v") {
        event.preventDefault();
        void pasteSelection(event.shiftKey ? "display" : "raw");
        return;
      }

      if (
        event.key === "Delete" &&
        !event.shiftKey &&
        !event.ctrlKey &&
        !event.metaKey
      ) {
        event.preventDefault();
        deleteSelection();
      }
    };

    window.addEventListener("keydown", handleWindowKeyDown, true);

    return () => {
      window.removeEventListener("keydown", handleWindowKeyDown, true);
    };
  }, [
    copySelection,
    cutSelection,
    deleteSelection,
    isChartEditorOpen,
    pasteSelection,
  ]);

  return (
    <main className="app-shell">
      <section className="formula-bar" aria-label="Formula bar">
        <div className="formula-bar__address">
          {selectedCellAddress || "Cell"}
        </div>
        <div className="formula-bar__field">
          <input
            aria-label="Selected cell formula or value"
            className="formula-bar__input"
            disabled={!selectedCell}
            id="formula-input"
            ref={formulaInputRef}
            onBlur={() => {
              void commitFormulaBar();
            }}
            onChange={handleFormulaInputChange}
            onKeyDown={handleFormulaInputKeyDown}
            placeholder={
              selectedCell
                ? "Type a value or a formula like =A1+B1"
                : "Select a cell to inspect or edit"
            }
            value={selectedCell ? formulaInputValue : ""}
          />
        </div>
      </section>

      <div className="app-shell__workspace">
        <section
          className="sheet-surface"
          aria-label="Spreadsheet surface"
          onPointerDown={() => {
            setSelectedChartId(null);
          }}
          ref={sheetSurfaceRef}
        >
          <DataEditor
            onCellContextMenu={handleCellContextMenu}
            columns={columns}
            getCellContent={getCellContent}
            getCellsForSelection={getCellsForSelection}
            ref={gridRef}
            gridSelection={gridSelection}
            height="100%"
            onCellEdited={handleCellEdited}
            onDelete={(selection) => {
              deleteSelection(selection);
              return false;
            }}
            onGridSelectionChange={setGridSelection}
            onPaste={handlePaste}
            onSelectionCleared={() => {
              setGridSelection(createEmptyGridSelection());
            }}
            onVisibleRegionChanged={(region) => {
              lastVisibleRegionRef.current = {
                height: region.height,
                width: region.width,
                x: region.x,
                y: region.y,
              };

              setGridViewportNonce((current) => current + 1);
              void loadVisibleRange(lastVisibleRegionRef.current);
            }}
            rowMarkers={{ kind: "number", startIndex: 1, width: 60 }}
            rows={rowCount}
            smoothScrollX
            smoothScrollY
            theme={GRID_THEME}
            width="100%"
          />
          <WorkbookChartOverlay
            gridRef={gridRef}
            isLoading={isSheetChartPreviewsLoading}
            onCommitChartLayout={commitChartLayout}
            onEditChart={openEditChartEditor}
            onSelectChart={setSelectedChartId}
            previews={sheetChartPreviews?.previews ?? []}
            selectedChartId={selectedChartId}
            surfaceRef={sheetSurfaceRef}
            viewportNonce={gridViewportNonce}
          />
        </section>
      </div>

      <footer className="app-shell__status-bar" aria-label="Workbook status">
        <div className="app-shell__meta">
          <label className="app-shell__selector">
            <select
              aria-label="Active sheet"
              className="app-shell__select"
              onChange={handleActiveSheetChange}
              value={activeSheet?.id ?? ""}
            >
              {(sheetSummary?.sheets ?? []).map((sheet) => (
                <option key={sheet.id} value={sheet.id}>
                  {sheet.name}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div className="app-shell__stats" aria-label="Workbook state">
          <span>{`${rowCount}x${columnCount}`}</span>
          <span>{`${activeSheetChartEntries.length} charts`}</span>
          <span>{sheetSummary ? `v${sheetSummary.version}` : "syncing"}</span>
        </div>
      </footer>

      {chartEditorSession ? (
        <ChartEditorDialog
          expectedVersion={chartEditorSession.expectedVersion}
          key={
            chartEditorSession.request.mode === "edit"
              ? `edit:${chartEditorSession.request.chartId}:${chartEditorSession.expectedVersion}`
              : `create:${chartEditorSession.request.sheetId ?? "active"}:${chartEditorSession.expectedVersion}`
          }
          onClose={closeChartEditor}
          onVersionConflict={handleChartEditorVersionConflict}
          request={chartEditorSession.request}
        />
      ) : null}

      <ToastViewport onDismiss={dismissToast} toasts={toasts} />
    </main>
  );
}
