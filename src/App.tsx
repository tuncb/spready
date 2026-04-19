import DataEditor, {
  CompactSelection,
  GridCellKind,
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

import { APP_MENU_ACTIONS, type AppMenuAction } from "./app-menu";
import {
  getColumnTitle,
  isFormulaInput,
  type CellDataResult,
  type SheetDisplayRangeResult,
  type SheetRangeRequest,
  type SheetRangeResult,
  type WorkbookSummary,
} from "./workbook-core";

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

function createLoadingCell(): GridCell {
  return {
    allowOverlay: false,
    kind: GridCellKind.Loading,
  };
}

function createTextCell(input: string, display: string): GridCell {
  return {
    allowOverlay: true,
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

function getDefaultWorkbookFilePath(summary: WorkbookSummary | null): string {
  return summary?.documentFilePath ?? DEFAULT_WORKBOOK_FILE_NAME;
}

export default function App() {
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [formulaInputValue, setFormulaInputValue] = useState("");
  const [gridSelection, setGridSelection] = useState<GridSelection>(
    createEmptyGridSelection,
  );
  const [selectedCellData, setSelectedCellData] =
    useState<CellDataResult | null>(null);
  const [sheetSummary, setSheetSummary] = useState<WorkbookSummary | null>(
    null,
  );
  const [viewNonce, setViewNonce] = useState(0);

  const exportPathRef = useRef<string>();
  const displayRangeCacheRef = useRef<SheetDisplayRangeResult | null>(null);
  const lastVisibleRegionRef = useRef<VisibleRegion | null>(null);
  const pendingCellDataRequestIdRef = useRef(0);
  const pendingRangeRequestIdRef = useRef(0);
  const rawRangeCacheRef = useRef<SheetRangeResult | null>(null);

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
  const rowCount = activeSheet?.rowCount ?? 1;
  const columnCount = activeSheet?.columnCount ?? 1;
  const columns = useMemo(() => createColumns(columnCount), [columnCount]);

  const applyTransaction = useCallback(
    async (
      operations: Parameters<
        typeof window.appShell.applyTransaction
      >[0]["operations"],
    ) => {
      const result = await window.appShell.applyTransaction({ operations });

      setSheetSummary(result.summary);
      setErrorMessage(null);

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
        setErrorMessage(getErrorMessage(error));
      }
    },
    [activeSheet],
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
      setErrorMessage(null);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    }
  }, [activeSheet, selectedCell]);

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
        setErrorMessage(getErrorMessage(error));
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
      });
    },
    [
      activeSheet,
      applyTransaction,
      loadVisibleRange,
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
        setErrorMessage(getErrorMessage(error));
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
      });

      return true;
    },
    [
      activeSheet,
      applyTransaction,
      loadVisibleRange,
      refreshSelectedCellData,
      selectedCell,
    ],
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
      setErrorMessage(getErrorMessage(error));
      void loadVisibleRange(lastVisibleRegionRef.current);
      void refreshSelectedCellData();
    }
  }, [
    activeSheet,
    applyTransaction,
    formulaInputValue,
    loadVisibleRange,
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
      setErrorMessage(getErrorMessage(error));
    });
  }, [activeSheet, applyTransaction]);

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
      setErrorMessage(getErrorMessage(error));
    });
  }, [activeSheet, applyTransaction]);

  const addSheet = useCallback(() => {
    void applyTransaction([
      {
        activate: true,
        type: "addSheet",
      },
    ]).catch((error) => {
      setErrorMessage(getErrorMessage(error));
    });
  }, [applyTransaction]);

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
      setErrorMessage(getErrorMessage(error));
    });
  }, [activeSheet, applyTransaction]);

  const handleActiveSheetChange = useCallback(
    (event: ChangeEvent<HTMLSelectElement>) => {
      void applyTransaction([
        {
          sheetId: event.target.value,
          type: "setActiveSheet",
        },
      ]).catch((error) => {
        setErrorMessage(getErrorMessage(error));
      });
    },
    [applyTransaction],
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
      setErrorMessage(getErrorMessage(error));
    }
  }, [applyTransaction]);

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
      setErrorMessage(getErrorMessage(error));
    }
  }, [activeSheet, applyTransaction]);

  const handleOpenWorkbook = useCallback(async () => {
    try {
      const result = await window.appShell.openWorkbookFile();

      if (result.canceled) {
        return;
      }

      setSheetSummary(result.summary);
      setErrorMessage(null);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    }
  }, []);

  const handleSaveWorkbookAs = useCallback(async () => {
    try {
      const result = await window.appShell.saveWorkbookFileAs(
        getDefaultWorkbookFilePath(sheetSummary),
      );

      if (result.canceled) {
        return;
      }

      setSheetSummary(result.summary);
      setErrorMessage(null);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    }
  }, [sheetSummary]);

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
      setErrorMessage(null);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    }
  }, [handleSaveWorkbookAs, sheetSummary]);

  useEffect(() => {
    let isMounted = true;

    void window.appShell
      .getWorkbookSummary()
      .then((summary) => {
        if (!isMounted) {
          return;
        }

        setSheetSummary(summary);
        setErrorMessage(null);
      })
      .catch((error) => {
        if (!isMounted) {
          return;
        }

        setErrorMessage(getErrorMessage(error));
      });

    const unsubscribeWorkbook = window.appShell.onWorkbookChanged((summary) => {
      setSheetSummary(summary);
      setErrorMessage(null);
    });

    return () => {
      isMounted = false;
      unsubscribeWorkbook();
    };
  }, []);

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
    return window.appShell.onMenuAction((action) => {
      const handleMenuAction = (nextAction: AppMenuAction) => {
        switch (nextAction) {
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
        }
      };

      handleMenuAction(action);
    });
  }, [
    addColumn,
    addRow,
    addSheet,
    deleteSheet,
    handleExport,
    handleImport,
    handleOpenWorkbook,
    handleSaveWorkbook,
    handleSaveWorkbookAs,
  ]);

  return (
    <main className="app-shell">
      {errorMessage ? (
        <div className="app-shell__error" role="status">
          {errorMessage}
        </div>
      ) : null}

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

      <section className="sheet-surface" aria-label="Spreadsheet surface">
        <DataEditor
          columns={columns}
          getCellContent={getCellContent}
          getCellsForSelection={getCellsForSelection}
          gridSelection={gridSelection}
          height="100%"
          onCellEdited={handleCellEdited}
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

            void loadVisibleRange(lastVisibleRegionRef.current);
          }}
          rowMarkers={{ kind: "number", startIndex: 1, width: 60 }}
          rows={rowCount}
          smoothScrollX
          smoothScrollY
          theme={GRID_THEME}
          width="100%"
        />
      </section>

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
          <span>{rowCount} rows</span>
          <span>{columnCount} columns</span>
          <span>{sheetSummary ? `v${sheetSummary.version}` : "syncing"}</span>
        </div>
      </footer>
    </main>
  );
}
