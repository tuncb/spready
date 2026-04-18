import DataEditor, {
  GridCellKind,
  type EditableGridCell,
  type GridCell,
  type GridColumn,
  type Item,
} from '@glideapps/glide-data-grid';
import { type ChangeEvent, useCallback, useEffect, useMemo, useRef, useState } from 'react';

import {
  getColumnTitle,
  type ControlServerInfo,
  type SheetRangeRequest,
  type SheetRangeResult,
  type WorkbookSummary,
} from './workbook-core';

const DEFAULT_COLUMN_WIDTH = 140;
const DEFAULT_VISIBLE_COLUMN_COUNT = 10;
const DEFAULT_VISIBLE_ROW_COUNT = 36;
const VISIBLE_COLUMN_PADDING = 4;
const VISIBLE_ROW_PADDING = 24;

type VisibleRegion = {
  height: number;
  width: number;
  x: number;
  y: number;
};

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
      Math.min(columnCount - startColumn, targetRegion.width + VISIBLE_COLUMN_PADDING * 2),
    ),
    rowCount: Math.max(
      1,
      Math.min(rowCount - startRow, targetRegion.height + VISIBLE_ROW_PADDING * 2),
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

function createLoadingCell(): GridCell {
  return {
    allowOverlay: false,
    kind: GridCellKind.Loading,
  };
}

function createTextCell(value: string): GridCell {
  return {
    allowOverlay: true,
    data: value,
    displayData: value,
    kind: GridCellKind.Text,
  };
}

function getCachedCellValue(
  cache: SheetRangeResult | null,
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

function setCachedCellValue(
  cache: SheetRangeResult | null,
  columnIndex: number,
  rowIndex: number,
  sheetId: string,
  value: string,
): SheetRangeResult | null {
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
  return error instanceof Error ? error.message : 'Unknown error';
}

export default function App() {
  const [controlInfo, setControlInfo] = useState<ControlServerInfo | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [sheetSummary, setSheetSummary] = useState<WorkbookSummary | null>(null);
  const [viewNonce, setViewNonce] = useState(0);

  const exportPathRef = useRef<string>();
  const lastVisibleRegionRef = useRef<VisibleRegion | null>(null);
  const pendingRangeRequestIdRef = useRef(0);
  const rangeCacheRef = useRef<SheetRangeResult | null>(null);

  const activeSheet = useMemo(
    () => sheetSummary?.sheets.find((sheet) => sheet.id === sheetSummary.activeSheetId) ?? null,
    [sheetSummary],
  );

  const rowCount = activeSheet?.rowCount ?? 1;
  const columnCount = activeSheet?.columnCount ?? 1;
  const columns = useMemo(() => createColumns(columnCount), [columnCount]);

  const applyTransaction = useCallback(
    async (operations: Parameters<typeof window.appShell.applyTransaction>[0]['operations']) => {
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

      const request = buildRangeRequest(activeSheet.id, activeSheet.columnCount, activeSheet.rowCount, region);
      const requestId = pendingRangeRequestIdRef.current + 1;

      pendingRangeRequestIdRef.current = requestId;

      try {
        const result = await window.appShell.getSheetRange(request);

        if (pendingRangeRequestIdRef.current !== requestId) {
          return;
        }

        rangeCacheRef.current = result;
        setViewNonce((current) => current + 1);
      } catch (error) {
        setErrorMessage(getErrorMessage(error));
      }
    },
    [activeSheet],
  );

  const getCellContent = useCallback(
    (cell: Item): GridCell => {
      const [columnIndex, rowIndex] = cell;
      const cachedValue = getCachedCellValue(
        rangeCacheRef.current,
        columnIndex,
        rowIndex,
        activeSheet?.id,
      );

      if (cachedValue === undefined) {
        return createLoadingCell();
      }

      return createTextCell(cachedValue);
    },
    [activeSheet?.id, viewNonce],
  );

  const getCellsForSelection = useCallback(
    (selection: VisibleRegion) => {
      return async () => {
        if (!activeSheet) {
          return [];
        }

        const range = await window.appShell.getSheetRange({
          columnCount: selection.width,
          rowCount: selection.height,
          sheetId: activeSheet.id,
          startColumn: selection.x,
          startRow: selection.y,
        });

        return range.values.map((row) => row.map(createTextCell));
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

      rangeCacheRef.current = setCachedCellValue(
        rangeCacheRef.current,
        columnIndex,
        rowIndex,
        activeSheet.id,
        newValue.data,
      );
      setViewNonce((current) => current + 1);

      void applyTransaction([
        {
          columnIndex,
          rowIndex,
          type: 'setCell',
          value: newValue.data,
        },
      ]).catch((error) => {
        setErrorMessage(getErrorMessage(error));
      });
    },
    [activeSheet, applyTransaction],
  );

  const handlePaste = useCallback(
    (target: Item, values: readonly (readonly string[])[]) => {
      if (!activeSheet || values.length === 0) {
        return false;
      }

      const [startColumn, startRow] = target;
      const nextValues = values.map((row) => [...row]);

      for (let rowOffset = 0; rowOffset < nextValues.length; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < nextValues[rowOffset].length; columnOffset += 1) {
          rangeCacheRef.current = setCachedCellValue(
            rangeCacheRef.current,
            startColumn + columnOffset,
            startRow + rowOffset,
            activeSheet.id,
            nextValues[rowOffset][columnOffset] ?? '',
          );
        }
      }

      setViewNonce((current) => current + 1);

      void applyTransaction([
        {
          startColumn,
          startRow,
          type: 'setRange',
          values: nextValues,
        },
      ]).catch((error) => {
        setErrorMessage(getErrorMessage(error));
      });

      return true;
    },
    [activeSheet, applyTransaction],
  );

  const addColumn = useCallback(() => {
    if (!activeSheet) {
      return;
    }

    void applyTransaction([
      {
        columnIndex: activeSheet.columnCount,
        count: 1,
        type: 'insertColumns',
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
        type: 'insertRows',
      },
    ]).catch((error) => {
      setErrorMessage(getErrorMessage(error));
    });
  }, [activeSheet, applyTransaction]);

  const addSheet = useCallback(() => {
    void applyTransaction([
      {
        activate: true,
        type: 'addSheet',
      },
    ]).catch((error) => {
      setErrorMessage(getErrorMessage(error));
    });
  }, [applyTransaction]);

  const handleActiveSheetChange = useCallback(
    (event: ChangeEvent<HTMLSelectElement>) => {
      void applyTransaction([
        {
          sheetId: event.target.value,
          type: 'setActiveSheet',
        },
      ]).catch((error) => {
        setErrorMessage(getErrorMessage(error));
      });
    },
    [applyTransaction],
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
          type: 'replaceSheetFromCsv',
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
        `${activeSheet.name.replaceAll(/\s+/g, '-') || 'Sheet'}.csv`;
      const result = await window.appShell.saveCsvFile(csv, defaultPath);

      if (result.canceled) {
        return;
      }

      exportPathRef.current = result.filePath;

      await applyTransaction([
        {
          sourceFilePath: result.filePath,
          type: 'setSheetSourceFile',
        },
      ]);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    }
  }, [activeSheet, applyTransaction]);

  useEffect(() => {
    let isMounted = true;

    void Promise.all([window.appShell.getWorkbookSummary(), window.appShell.getControlInfo()])
      .then(([summary, nextControlInfo]) => {
        if (!isMounted) {
          return;
        }

        setSheetSummary(summary);
        setControlInfo(nextControlInfo);
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
    if (!activeSheet) {
      return;
    }

    rangeCacheRef.current = null;
    setViewNonce((current) => current + 1);

    if (activeSheet.sourceFilePath) {
      exportPathRef.current = activeSheet.sourceFilePath;
    }

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
      if (action === 'import') {
        void handleImport();
        return;
      }

      void handleExport();
    });
  }, [handleExport, handleImport]);

  return (
    <main className="app-shell">
      <header className="app-shell__toolbar">
        <div className="app-shell__brand">
          <div className="app-shell__meta">
            <label className="app-shell__selector">
              <span>Active Sheet</span>
              <select
                className="app-shell__select"
                onChange={handleActiveSheetChange}
                value={activeSheet?.id ?? ''}
              >
                {(sheetSummary?.sheets ?? []).map((sheet) => (
                  <option key={sheet.id} value={sheet.id}>
                    {sheet.name}
                  </option>
                ))}
              </select>
            </label>

            {controlInfo ? (
              <span className="app-shell__control">
                tcp://{controlInfo.host}:{controlInfo.port}
              </span>
            ) : null}
          </div>
        </div>

        <div className="app-shell__stats" aria-label="Workbook state">
          <span>{rowCount} rows</span>
          <span>{columnCount} columns</span>
          <span>{sheetSummary?.sheets.length ?? 0} sheets</span>
          <span>{sheetSummary ? `v${sheetSummary.version}` : 'syncing'}</span>
        </div>

        <div className="app-shell__actions">
          <button className="app-shell__button" type="button" onClick={addRow}>
            Add Row
          </button>
          <button
            className="app-shell__button app-shell__button--secondary"
            type="button"
            onClick={addColumn}
          >
            Add Column
          </button>
          <button
            className="app-shell__button app-shell__button--secondary"
            type="button"
            onClick={addSheet}
          >
            New Sheet
          </button>
        </div>
      </header>

      {errorMessage ? (
        <div className="app-shell__status" role="status">
          {errorMessage}
        </div>
      ) : null}

      <section className="sheet-surface" aria-label="Spreadsheet surface">
        <DataEditor
          columns={columns}
          getCellContent={getCellContent}
          getCellsForSelection={getCellsForSelection}
          height="100%"
          onCellEdited={handleCellEdited}
          onPaste={handlePaste}
          onVisibleRegionChanged={(region) => {
            lastVisibleRegionRef.current = {
              height: region.height,
              width: region.width,
              x: region.x,
              y: region.y,
            };

            void loadVisibleRange(lastVisibleRegionRef.current);
          }}
          rowMarkers={{ kind: 'number', startIndex: 1, width: 60 }}
          rows={rowCount}
          smoothScrollX
          smoothScrollY
          width="100%"
        />
      </section>
    </main>
  );
}
