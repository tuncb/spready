import DataEditor, {
  CompactSelection,
  GridCellKind,
  TextCellEntry,
  type DataEditorProps,
  type DataEditorRef,
  type EditableGridCell,
  type GridCell,
  type GridColumn,
  type GridSelection,
  type Item,
  type TextCell,
  type Theme,
} from "@glideapps/glide-data-grid";
import {
  type ChangeEvent,
  type FormEvent,
  type KeyboardEvent,
  lazy,
  Suspense,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { flushSync } from "react-dom";

import { APP_MENU_ACTIONS, type AppMenuAction } from "./app-menu";
import {
  getCellEditArrowKeyMovement,
  isEditableShortcutTarget,
  shouldCloseWorkbookSearchOnEscape,
} from "./app-keyboard";
import { AboutDialog } from "./AboutDialog";
import { CellFormatDialog } from "./CellFormatDialog";
import { type ChartEditorWindowRequest } from "./chart-editor-state";
import { ChartEditorDialog } from "./ChartEditorWindow";
import { isDialogBackdropClick } from "./dialog-events";
import { InstallationDialog } from "./InstallationDialog";
import { RenameSheetDialog } from "./RenameSheetDialog";
import {
  createSortTableOperation,
  getNextTableSortDirection,
  getVisibleTableHeaderSortTargets,
  type TableHeaderSortTarget,
} from "./app-table-sort-controls";
import {
  buildStickyTableHeaderRangeRequest,
  getStickyTableHeader,
  getTableHeaderCacheKey,
} from "./app-table-headers";
import { getTableColumnHighlightStyle } from "./app-table-column-highlights";
import {
  buildTableRowHintItems,
  buildTableRowHintRangeRequests,
  getTableRowHintTarget,
  getTableRowHintTableKey,
  getTableRowHintTargetKey,
  type TableRowHintItem,
  type TableRowHintTarget,
} from "./app-table-row-hints";
import {
  DEFAULT_COLUMN_WIDTH,
  getColumnTitle,
  isFormulaInput,
  MAX_COLUMN_WIDTH,
  MIN_COLUMN_WIDTH,
  parseTsv,
  type CellDataResult,
  type ClipboardRangePayload,
  type ClipboardRangeMode,
  type WorkbookCellStyle,
  type WorkbookChartLayout,
  type WorkbookTableSummary,
  type WorkbookSheetChartPreviewsResult,
  type SheetSummary,
  type SheetDisplayRangeResult,
  type SheetRangeRequest,
  type SheetRangeResult,
  type SheetStyleRangeResult,
  type WorkbookSearchResult,
  type WorkbookSearchScope,
  type WorkbookSearchState,
  type WorkbookSearchValueMode,
  type WorkbookSummary,
  type WorkbookTransactionOperation,
} from "./workbook-core";
import { ToastViewport } from "./ToastViewport";
import { enqueueToast, removeToast, type ToastNotification } from "./toast-state";
import { filterSheetQuickOpenResults } from "./sheet-quick-open";

const LazyWorkbookChartOverlay = lazy(() =>
  import("./WorkbookChartOverlay").then((module) => ({
    default: module.WorkbookChartOverlay,
  })),
);

const DEFAULT_CELL_FONT_SIZE = 13;
const DEFAULT_VISIBLE_COLUMN_COUNT = 10;
const DEFAULT_VISIBLE_ROW_COUNT = 36;
const DEFAULT_WORKBOOK_FILE_NAME = "Workbook.spready";
const TABLE_SORT_BUTTON_SIZE = 16.5;
const TABLE_SORT_CONTROL_WIDTH = TABLE_SORT_BUTTON_SIZE;
const VISIBLE_COLUMN_PADDING = 4;
const VISIBLE_ROW_PADDING = 24;

let didLogAppRenderStart = false;
let didLogInitialChartPreviewsDone = false;
let didLogInitialChartPreviewsStart = false;
let didLogInitialRangeDone = false;
let didLogInitialRangeStart = false;
let didLogInitialWorkbookSummaryDone = false;
let didLogInitialWorkbookSummaryStart = false;

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

const SEARCH_MATCH_HIGHLIGHT_COLOR = "rgba(245, 158, 11, 0.24)";
const SEARCH_ACTIVE_CELL_THEME: Partial<Theme> = {
  accentColor: "#b45309",
  accentLight: "rgba(245, 158, 11, 0.28)",
  bgCell: "#fef3c7",
  textDark: "#78350f",
};
const DEFAULT_SEARCH_STATE: WorkbookSearchState = {
  activeResult: null,
  query: {
    activeResultIndex: -1,
    caseSensitive: false,
    scope: "sheet",
    text: "",
    valueMode: "display",
    wholeWord: false,
  },
  results: [],
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

type CellFormatSession = {
  expectedVersion: number;
  initialStyle?: WorkbookCellStyle;
  ranges: SheetRangeRequest[];
};

type RenameSheetSession = {
  expectedVersion: number;
  name: string;
  sheetId: string;
};

type InstallationDialogMode = "check-updates" | "manage";

type RangeCache = SheetDisplayRangeResult | SheetRangeResult;
type SearchHighlightRegions = NonNullable<DataEditorProps["highlightRegions"]>;

interface WorkbookSearchBoxProps {
  onChangeQuery: (
    text: string,
    scope: WorkbookSearchScope,
    valueMode: WorkbookSearchValueMode,
    caseSensitive: boolean,
    wholeWord: boolean,
  ) => void;
  onClose: () => void;
  onNext: () => void;
  onPrevious: () => void;
  state: WorkbookSearchState;
}

function WorkbookSearchBox({
  onChangeQuery,
  onClose,
  onNext,
  onPrevious,
  state,
}: WorkbookSearchBoxProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const resultLabel =
    state.query.text.length === 0
      ? ""
      : state.results.length === 0
        ? "0 / 0"
        : `${state.query.activeResultIndex + 1} / ${state.results.length}`;

  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.select();
  }, []);

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    onNext();
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLFormElement>) => {
    if (event.key === "Enter" && event.shiftKey) {
      event.preventDefault();
      onPrevious();
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      onClose();
    }
  };

  return (
    <form
      aria-label="Search workbook"
      className="workbook-search"
      onKeyDown={handleKeyDown}
      onSubmit={handleSubmit}
    >
      <input
        aria-label="Search cells"
        autoComplete="off"
        className="workbook-search__input"
        onChange={(event) => {
          onChangeQuery(
            event.target.value,
            state.query.scope,
            state.query.valueMode,
            state.query.caseSensitive,
            state.query.wholeWord,
          );
        }}
        placeholder="Find"
        ref={inputRef}
        value={state.query.text}
      />
      <div className="workbook-search__options" role="group" aria-label="Search match options">
        <button
          aria-label="Case sensitive search"
          aria-pressed={state.query.caseSensitive}
          className={
            state.query.caseSensitive
              ? "workbook-search__option-button is-active"
              : "workbook-search__option-button"
          }
          onClick={() => {
            onChangeQuery(
              state.query.text,
              state.query.scope,
              state.query.valueMode,
              !state.query.caseSensitive,
              state.query.wholeWord,
            );
          }}
          title="Case sensitive"
          type="button"
        >
          Aa
        </button>
        <button
          aria-label="Whole word search"
          aria-pressed={state.query.wholeWord}
          className={
            state.query.wholeWord
              ? "workbook-search__option-button is-active"
              : "workbook-search__option-button"
          }
          onClick={() => {
            onChangeQuery(
              state.query.text,
              state.query.scope,
              state.query.valueMode,
              state.query.caseSensitive,
              !state.query.wholeWord,
            );
          }}
          title="Whole word"
          type="button"
        >
          W
        </button>
      </div>
      <div className="workbook-search__scope" role="group" aria-label="Search scope">
        {(["sheet", "workbook"] as const).map((scope) => (
          <button
            aria-pressed={state.query.scope === scope}
            className={
              state.query.scope === scope
                ? "workbook-search__scope-button is-active"
                : "workbook-search__scope-button"
            }
            key={scope}
            onClick={() => {
              onChangeQuery(
                state.query.text,
                scope,
                state.query.valueMode,
                state.query.caseSensitive,
                state.query.wholeWord,
              );
            }}
            type="button"
          >
            {scope === "sheet" ? "Sheet" : "Workbook"}
          </button>
        ))}
      </div>
      <div className="workbook-search__mode" role="group" aria-label="Search values">
        {(["display", "raw"] as const).map((valueMode) => (
          <button
            aria-pressed={state.query.valueMode === valueMode}
            className={
              state.query.valueMode === valueMode
                ? "workbook-search__mode-button is-active"
                : "workbook-search__mode-button"
            }
            key={valueMode}
            onClick={() => {
              onChangeQuery(
                state.query.text,
                state.query.scope,
                valueMode,
                state.query.caseSensitive,
                state.query.wholeWord,
              );
            }}
            type="button"
          >
            {valueMode === "display" ? "Displayed" : "Raw"}
          </button>
        ))}
      </div>
      <span className="workbook-search__count">{resultLabel}</span>
      <button
        aria-label="Previous search result"
        className="workbook-search__button"
        disabled={state.results.length === 0}
        onClick={onPrevious}
        type="button"
      >
        Prev
      </button>
      <button
        aria-label="Next search result"
        className="workbook-search__button"
        disabled={state.results.length === 0}
        type="submit"
      >
        Next
      </button>
      <button
        aria-label="Close search"
        className="workbook-search__close"
        onClick={onClose}
        type="button"
      >
        Close
      </button>
    </form>
  );
}

interface SheetQuickOpenProps {
  activeSheetId: string | null;
  onClose: () => void;
  onSelect: (sheetId: string) => void;
  sheets: readonly SheetSummary[];
}

function SheetQuickOpen({ activeSheetId, onClose, onSelect, sheets }: SheetQuickOpenProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const [highlightedSheetId, setHighlightedSheetId] = useState<string | null>(activeSheetId);
  const [query, setQuery] = useState("");
  const filteredSheets = useMemo(() => filterSheetQuickOpenResults(sheets, query), [query, sheets]);
  const trimmedQuery = query.trim();
  const suggestedSheetId = useMemo(() => {
    if (filteredSheets.length === 0) {
      return null;
    }

    if (!trimmedQuery) {
      return filteredSheets.find((sheet) => sheet.id === activeSheetId)?.id ?? filteredSheets[0].id;
    }

    return filteredSheets[0].id;
  }, [activeSheetId, filteredSheets, trimmedQuery]);
  const highlightedSheet =
    filteredSheets.find((sheet) => sheet.id === highlightedSheetId) ?? filteredSheets[0] ?? null;

  useEffect(() => {
    const dialog = dialogRef.current;

    if (!dialog || dialog.open) {
      return;
    }

    dialog.showModal();
    inputRef.current?.focus();

    return () => {
      if (dialog.open) {
        dialog.close();
      }
    };
  }, []);

  useEffect(() => {
    setHighlightedSheetId(suggestedSheetId);
  }, [suggestedSheetId]);

  const moveHighlight = (direction: 1 | -1) => {
    if (filteredSheets.length === 0) {
      return;
    }

    const currentIndex = Math.max(
      0,
      filteredSheets.findIndex((sheet) => sheet.id === highlightedSheetId),
    );
    const nextIndex = (currentIndex + direction + filteredSheets.length) % filteredSheets.length;

    setHighlightedSheetId(filteredSheets[nextIndex].id);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLFormElement>) => {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      moveHighlight(1);
      return;
    }

    if (event.key === "ArrowUp") {
      event.preventDefault();
      moveHighlight(-1);
    }
  };

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!highlightedSheet) {
      return;
    }

    onSelect(highlightedSheet.id);
  };

  return (
    <dialog
      aria-label="Go to sheet"
      className="sheet-quick-open-dialog"
      onCancel={(event) => {
        event.preventDefault();
        onClose();
      }}
      onClick={(event) => {
        if (isDialogBackdropClick(event)) {
          onClose();
        }
      }}
      ref={dialogRef}
    >
      <form className="sheet-quick-open" onKeyDown={handleKeyDown} onSubmit={handleSubmit}>
        <label className="sheet-quick-open__label" htmlFor="sheet-quick-open-input">
          Go to sheet
        </label>
        <input
          aria-activedescendant={
            highlightedSheet ? getSheetQuickOpenOptionId(highlightedSheet.id) : undefined
          }
          aria-controls="sheet-quick-open-results"
          aria-expanded="true"
          aria-label="Search sheets"
          autoComplete="off"
          className="sheet-quick-open__input"
          id="sheet-quick-open-input"
          onChange={(event) => {
            setQuery(event.target.value);
          }}
          placeholder="Search sheets"
          ref={inputRef}
          role="combobox"
          value={query}
        />
        <ul className="sheet-quick-open__results" id="sheet-quick-open-results" role="listbox">
          {filteredSheets.length > 0 ? (
            filteredSheets.map((sheet) => {
              const isHighlighted = sheet.id === highlightedSheet?.id;

              return (
                <li
                  aria-selected={isHighlighted}
                  className={
                    isHighlighted
                      ? "sheet-quick-open__result is-highlighted"
                      : "sheet-quick-open__result"
                  }
                  id={getSheetQuickOpenOptionId(sheet.id)}
                  key={sheet.id}
                  role="option"
                >
                  <button
                    className="sheet-quick-open__result-button"
                    onClick={() => {
                      onSelect(sheet.id);
                    }}
                    onMouseDown={(event) => {
                      event.preventDefault();
                    }}
                    type="button"
                  >
                    <span className="sheet-quick-open__result-name">{sheet.name}</span>
                    <span className="sheet-quick-open__result-meta">
                      {`${sheet.rowCount}x${sheet.columnCount}`}
                    </span>
                  </button>
                </li>
              );
            })
          ) : (
            <li className="sheet-quick-open__empty">No sheets found</li>
          )}
        </ul>
      </form>
    </dialog>
  );
}

function getSheetQuickOpenOptionId(sheetId: string): string {
  return `sheet-quick-open-option-${sheetId.replaceAll(/[^A-Za-z0-9_-]/g, "_")}`;
}

type StyleRangeCache = SheetStyleRangeResult;
type TableSortControlPlacement = TableHeaderSortTarget & {
  height: number;
  left: number;
  top: number;
  width: number;
};

interface TableRowHintState {
  items: TableRowHintItem[];
  tableKey: string;
  tableName: string;
  tableRowNumber: number;
}

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

interface StickyTableHeaderColumns {
  headerValues: Readonly<Record<string, string>>;
  table: WorkbookTableSummary;
}

function createColumns(
  columnCount: number,
  columnWidths: Record<string, number>,
  stickyHeader?: StickyTableHeaderColumns | null,
): GridColumn[] {
  return Array.from({ length: columnCount }, (_, index) => ({
    id: `column-${index}`,
    ...(getStickyColumnTitle(index, stickyHeader) ?? { title: getColumnTitle(index) }),
    width: columnWidths[String(index)] ?? DEFAULT_COLUMN_WIDTH,
  }));
}

function getStickyColumnTitle(
  columnIndex: number,
  stickyHeader: StickyTableHeaderColumns | null | undefined,
): Pick<GridColumn, "themeOverride" | "title"> | null {
  if (!stickyHeader) {
    return null;
  }

  const tableStartColumn = stickyHeader.table.range.startColumn;
  const tableEndColumn =
    stickyHeader.table.range.startColumn + stickyHeader.table.range.columnCount;

  if (columnIndex < tableStartColumn || columnIndex >= tableEndColumn) {
    return null;
  }

  const headerTitle = stickyHeader.headerValues[String(columnIndex)]?.trim();

  if (!headerTitle) {
    return null;
  }

  return {
    themeOverride: {
      bgHeader: "#e8f5f2",
      bgHeaderHasFocus: "#d6eee8",
      bgHeaderHovered: "#ddf2ed",
      textHeader: "#0f3b34",
      textHeaderSelected: "#092c27",
    },
    title: headerTitle,
  };
}

function removeColumnResizeOverride(
  overrides: Record<string, number>,
  columnIndex: number,
  width: number,
): Record<string, number> {
  const key = String(columnIndex);

  if (overrides[key] !== width) {
    return overrides;
  }

  const nextOverrides = { ...overrides };

  delete nextOverrides[key];
  return nextOverrides;
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

function getCellThemeOverride(style: WorkbookCellStyle | undefined): Partial<Theme> | undefined {
  if (!style) {
    return undefined;
  }

  const fontParts = [
    style.italic ? "italic" : undefined,
    style.bold ? "700" : undefined,
    style.bold || style.italic || style.fontSize
      ? `${style.fontSize ?? DEFAULT_CELL_FONT_SIZE}px`
      : undefined,
  ].filter(Boolean);
  const theme: Partial<Theme> = {
    ...(style.backgroundColor ? { bgCell: style.backgroundColor } : {}),
    ...(style.fontFamily ? { fontFamily: style.fontFamily } : {}),
    ...(fontParts.length > 0 ? { baseFontStyle: fontParts.join(" ") } : {}),
    ...(style.textColor ? { textDark: style.textColor } : {}),
  };

  return Object.keys(theme).length > 0 ? theme : undefined;
}

function mergeCellThemeOverrides(
  baseTheme: Partial<Theme> | undefined,
  overlayTheme: Partial<Theme> | undefined,
): Partial<Theme> | undefined {
  if (!baseTheme && !overlayTheme) {
    return undefined;
  }

  return {
    ...(baseTheme ?? {}),
    ...(overlayTheme ?? {}),
  };
}

function createTextCell(
  input: string,
  display: string,
  style?: WorkbookCellStyle,
  themeOverride?: Partial<Theme>,
): GridCell {
  return {
    allowOverlay: true,
    allowWrapping: style?.wrapText,
    contentAlign: style?.horizontalAlign,
    copyData: input,
    data: input,
    displayData: display,
    kind: GridCellKind.Text,
    themeOverride: mergeCellThemeOverrides(getCellThemeOverride(style), themeOverride),
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

function getCachedCellStyle(
  cache: StyleRangeCache | null,
  columnIndex: number,
  rowIndex: number,
  sheetId?: string,
): WorkbookCellStyle | undefined {
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

  return cache.styles[rowOffset]?.[columnOffset] ?? undefined;
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

function isCellWithinSheetBounds(cell: Item, sheet: SheetSummary): boolean {
  const [columnIndex, rowIndex] = cell;

  return (
    Number.isInteger(columnIndex) &&
    Number.isInteger(rowIndex) &&
    columnIndex >= 0 &&
    rowIndex >= 0 &&
    columnIndex < sheet.columnCount &&
    rowIndex < sheet.rowCount
  );
}

function getSearchResultKey(result: WorkbookSearchResult): string {
  return `${result.sheetId}:${result.rowIndex}:${result.columnIndex}`;
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

function getSelectedStyleRanges(
  selection: GridSelection,
  sheetId: string,
  rowCount: number,
  columnCount: number,
): SheetRangeRequest[] {
  const ranges: SheetRangeRequest[] = [];
  const range = selection.current?.range;

  if (range) {
    ranges.push({
      columnCount: Math.max(1, range.width),
      rowCount: Math.max(1, range.height),
      sheetId,
      startColumn: range.x,
      startRow: range.y,
    });
  }

  for (const rowRun of compactSelectionToRuns(selection.rows)) {
    ranges.push({
      columnCount,
      rowCount: rowRun.count,
      sheetId,
      startColumn: 0,
      startRow: rowRun.start,
    });
  }

  for (const columnRun of compactSelectionToRuns(selection.columns)) {
    ranges.push({
      columnCount: columnRun.count,
      rowCount,
      sheetId,
      startColumn: columnRun.start,
      startRow: 0,
    });
  }

  return ranges;
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

function tableContainsCell(table: WorkbookTableSummary, cell: Item | null): boolean {
  if (!cell) {
    return false;
  }

  const [columnIndex, rowIndex] = cell;
  const range = table.range;

  return (
    rowIndex >= range.startRow &&
    rowIndex < range.startRow + range.rowCount &&
    columnIndex >= range.startColumn &&
    columnIndex < range.startColumn + range.columnCount
  );
}

function getTableContainingCell(
  tables: readonly WorkbookTableSummary[],
  cell: Item | null,
): WorkbookTableSummary | null {
  return tables.find((table) => tableContainsCell(table, cell)) ?? null;
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

function cellWouldBeCleared(selection: GridSelection, cell: Item | null): boolean {
  if (!cell) {
    return false;
  }

  if (selectionContainsCell(selection, cell)) {
    return true;
  }

  return selection.columns.hasIndex(cell[0]) || selection.rows.hasIndex(cell[1]);
}

function chartLayoutsEqual(left: WorkbookChartLayout, right: WorkbookChartLayout): boolean {
  return (
    left.height === right.height &&
    left.offsetX === right.offsetX &&
    left.offsetY === right.offsetY &&
    left.startColumn === right.startColumn &&
    left.startRow === right.startRow &&
    left.width === right.width &&
    left.zIndex === right.zIndex
  );
}

function replaceInputSelection(
  input: HTMLInputElement,
  nextText: string,
): { selectionStart: number; value: string } {
  const selectionStart = input.selectionStart ?? input.value.length;
  const selectionEnd = input.selectionEnd ?? input.value.length;
  const value = input.value.slice(0, selectionStart) + nextText + input.value.slice(selectionEnd);

  return {
    selectionStart: selectionStart + nextText.length,
    value,
  };
}

function getDefaultWorkbookFilePath(summary: WorkbookSummary | null): string {
  return summary?.documentFilePath ?? DEFAULT_WORKBOOK_FILE_NAME;
}

export default function App() {
  if (!didLogAppRenderStart) {
    didLogAppRenderStart = true;
    window.appShell.logStartupTiming("app-render-start");
  }

  const [cellFormatSession, setCellFormatSession] = useState<CellFormatSession | null>(null);
  const [installationDialogMode, setInstallationDialogMode] =
    useState<InstallationDialogMode | null>(null);
  const [chartEditorSession, setChartEditorSession] = useState<ChartEditorSession | null>(null);
  const [columnResizeOverrides, setColumnResizeOverrides] = useState<Record<string, number>>({});
  const [formulaInputValue, setFormulaInputValue] = useState("");
  const [gridSelection, setGridSelection] = useState<GridSelection>(createEmptyGridSelection);
  const [gridViewportNonce, setGridViewportNonce] = useState(0);
  const [isSheetChartPreviewsLoading, setIsSheetChartPreviewsLoading] = useState(false);
  const [isAboutDialogOpen, setIsAboutDialogOpen] = useState(false);
  const [isSheetQuickOpenOpen, setIsSheetQuickOpenOpen] = useState(false);
  const [isSearchOpen, setIsSearchOpen] = useState(false);
  const [renameSheetSession, setRenameSheetSession] = useState<RenameSheetSession | null>(null);
  const [searchState, setSearchState] = useState<WorkbookSearchState>(DEFAULT_SEARCH_STATE);
  const [selectedChartId, setSelectedChartId] = useState<string | null>(null);
  const [selectedCellData, setSelectedCellData] = useState<CellDataResult | null>(null);
  const [sheetChartPreviews, setSheetChartPreviews] =
    useState<WorkbookSheetChartPreviewsResult | null>(null);
  const [sheetSummary, setSheetSummary] = useState<WorkbookSummary | null>(null);
  const [tableHeaderCache, setTableHeaderCache] = useState<Record<string, Record<string, string>>>(
    {},
  );
  const [tableRowHint, setTableRowHint] = useState<TableRowHintState | null>(null);
  const [tableSortControls, setTableSortControls] = useState<TableSortControlPlacement[]>([]);
  const [toasts, setToasts] = useState<ToastNotification[]>([]);
  const [viewNonce, setViewNonce] = useState(0);

  const exportPathRef = useRef<string>();
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const formulaRowContextItemsRef = useRef<HTMLDListElement>(null);
  const formulaRowContextScrollLeftRef = useRef(0);
  const displayRangeCacheRef = useRef<SheetDisplayRangeResult | null>(null);
  const gridRef = useRef<DataEditorRef>(null);
  const lastVisibleRegionRef = useRef<VisibleRegion | null>(null);
  const pendingCellDataRequestIdRef = useRef(0);
  const pendingRangeRequestIdRef = useRef(0);
  const pendingTableRowHintRequestIdRef = useRef(0);
  const pendingSearchResultRef = useRef<WorkbookSearchResult | null>(null);
  const pendingSheetChartPreviewsRequestIdRef = useRef(0);
  const rawRangeCacheRef = useRef<SheetRangeResult | null>(null);
  const sheetSurfaceRef = useRef<HTMLElement>(null);
  const styleRangeCacheRef = useRef<SheetStyleRangeResult | null>(null);
  const isChartEditorOpen = chartEditorSession !== null;
  const isCellFormatOpen = cellFormatSession !== null;
  const isAboutOpen = isAboutDialogOpen;
  const isInstallationDialogOpen = installationDialogMode !== null;
  const isRenameSheetOpen = renameSheetSession !== null;
  const isSheetQuickOpenDialogOpen = isSheetQuickOpenOpen;
  const isModalDialogOpen =
    isAboutOpen ||
    isChartEditorOpen ||
    isCellFormatOpen ||
    isInstallationDialogOpen ||
    isRenameSheetOpen ||
    isSheetQuickOpenDialogOpen;

  const activeSheet = useMemo(
    () => sheetSummary?.sheets.find((sheet) => sheet.id === sheetSummary.activeSheetId) ?? null,
    [sheetSummary],
  );
  const selectedCell = gridSelection.current?.cell ?? null;
  const selectedCellAddress = useMemo(() => getSelectedCellAddress(selectedCell), [selectedCell]);
  const activeSheetTableEntries = useMemo(
    () => (sheetSummary?.tables ?? []).filter((table) => table.sheetId === activeSheet?.id),
    [activeSheet?.id, sheetSummary?.tables],
  );
  const selectedTable = useMemo(
    () => getTableContainingCell(activeSheetTableEntries, selectedCell),
    [activeSheetTableEntries, selectedCell],
  );
  const tableRowHintTarget = useMemo<TableRowHintTarget | null>(
    () => getTableRowHintTarget(activeSheetTableEntries, selectedCell),
    [activeSheetTableEntries, selectedCell],
  );
  const tableRowHintTargetKey =
    tableRowHintTarget && sheetSummary
      ? getTableRowHintTargetKey(tableRowHintTarget, sheetSummary.version)
      : null;
  const rowCount = activeSheet?.rowCount ?? 1;
  const columnCount = activeSheet?.columnCount ?? 1;
  const effectiveColumnWidths = useMemo(
    () => ({
      ...(activeSheet?.columnWidths ?? {}),
      ...columnResizeOverrides,
    }),
    [activeSheet?.columnWidths, columnResizeOverrides],
  );
  const stickyTableHeader = useMemo(
    () => getStickyTableHeader(activeSheetTableEntries, lastVisibleRegionRef.current),
    [activeSheetTableEntries, gridViewportNonce],
  );
  const stickyTableHeaderRequest = useMemo(
    () =>
      stickyTableHeader
        ? buildStickyTableHeaderRangeRequest(stickyTableHeader, lastVisibleRegionRef.current)
        : null,
    [gridViewportNonce, stickyTableHeader],
  );
  const stickyTableHeaderCacheKey =
    activeSheet && stickyTableHeader && sheetSummary
      ? getTableHeaderCacheKey(activeSheet.id, stickyTableHeader.id, sheetSummary.version)
      : null;
  const stickyTableHeaderColumns = useMemo<StickyTableHeaderColumns | null>(() => {
    if (!stickyTableHeader || !stickyTableHeaderCacheKey) {
      return null;
    }

    return {
      headerValues: tableHeaderCache[stickyTableHeaderCacheKey] ?? {},
      table: stickyTableHeader,
    };
  }, [stickyTableHeader, stickyTableHeaderCacheKey, tableHeaderCache]);
  const columns = useMemo(
    () => createColumns(columnCount, effectiveColumnWidths, stickyTableHeaderColumns),
    [columnCount, effectiveColumnWidths, stickyTableHeaderColumns],
  );
  const currentSelectionRange = useMemo(
    () => (activeSheet ? getCurrentSelectionRange(gridSelection, activeSheet.id) : null),
    [activeSheet, gridSelection],
  );
  const activeSearchResultKey = useMemo(() => {
    const result = searchState.activeResult;

    if (!result || result.sheetId !== activeSheet?.id) {
      return null;
    }

    return getSearchResultKey(result);
  }, [activeSheet?.id, searchState.activeResult]);
  const searchHighlightRegions = useMemo<SearchHighlightRegions>(() => {
    if (!activeSheet || searchState.query.text.length === 0) {
      return [];
    }

    return searchState.results
      .filter(
        (result) =>
          result.sheetId === activeSheet.id && getSearchResultKey(result) !== activeSearchResultKey,
      )
      .map((result) => ({
        color: SEARCH_MATCH_HIGHLIGHT_COLOR,
        range: {
          height: 1,
          width: 1,
          x: result.columnIndex,
          y: result.rowIndex,
        },
      }));
  }, [activeSearchResultKey, activeSheet, searchState.query.text.length, searchState.results]);
  const shouldRenderChartOverlay = (sheetChartPreviews?.previews.length ?? 0) > 0;
  useLayoutEffect(() => {
    const animationFrameId = window.requestAnimationFrame(() => {
      const grid = gridRef.current;
      const surface = sheetSurfaceRef.current;

      if (!grid || !surface) {
        setTableSortControls([]);
        return;
      }

      const fallbackVisibleRegion =
        activeSheet === null
          ? null
          : {
              height: Math.min(activeSheet.rowCount, DEFAULT_VISIBLE_ROW_COUNT),
              width: Math.min(activeSheet.columnCount, DEFAULT_VISIBLE_COLUMN_COUNT),
              x: 0,
              y: 0,
            };
      const targets = getVisibleTableHeaderSortTargets(
        activeSheetTableEntries,
        lastVisibleRegionRef.current ?? fallbackVisibleRegion,
      );
      const surfaceBounds = surface.getBoundingClientRect();
      const placements = targets.flatMap<TableSortControlPlacement>((target) => {
        const cellBounds = grid.getBounds(target.columnIndex, target.rowIndex);

        if (!cellBounds || cellBounds.width < TABLE_SORT_CONTROL_WIDTH + 8) {
          return [];
        }

        return [
          {
            ...target,
            height: TABLE_SORT_BUTTON_SIZE,
            left:
              cellBounds.x -
              surfaceBounds.left +
              Math.max(0, cellBounds.width - TABLE_SORT_CONTROL_WIDTH - 6),
            top:
              cellBounds.y -
              surfaceBounds.top +
              Math.max(0, (cellBounds.height - TABLE_SORT_BUTTON_SIZE) / 2),
            width: TABLE_SORT_CONTROL_WIDTH,
          },
        ];
      });

      setTableSortControls(placements);
    });

    return () => {
      window.cancelAnimationFrame(animationFrameId);
    };
  }, [activeSheet, activeSheetTableEntries, effectiveColumnWidths, gridViewportNonce]);

  useLayoutEffect(() => {
    const items = formulaRowContextItemsRef.current;

    if (!items || !tableRowHint) {
      return;
    }

    items.scrollLeft = formulaRowContextScrollLeftRef.current;
  }, [tableRowHint]);

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

  useEffect(() => {
    if (!stickyTableHeaderCacheKey || !stickyTableHeaderRequest) {
      return;
    }

    const cachedValues = tableHeaderCache[stickyTableHeaderCacheKey] ?? {};
    let hasMissingHeader = false;

    for (
      let columnIndex = stickyTableHeaderRequest.startColumn;
      columnIndex < stickyTableHeaderRequest.startColumn + stickyTableHeaderRequest.columnCount;
      columnIndex += 1
    ) {
      if (cachedValues[String(columnIndex)] === undefined) {
        hasMissingHeader = true;
        break;
      }
    }

    if (!hasMissingHeader) {
      return;
    }

    let isCancelled = false;

    void window.appShell
      .getSheetDisplayRange(stickyTableHeaderRequest)
      .then((headerRange) => {
        if (isCancelled) {
          return;
        }

        setTableHeaderCache((current) => {
          const nextValues = { ...(current[stickyTableHeaderCacheKey] ?? {}) };
          const headerRow = headerRange.values[0] ?? [];

          for (let columnOffset = 0; columnOffset < headerRange.columnCount; columnOffset += 1) {
            nextValues[String(headerRange.startColumn + columnOffset)] =
              headerRow[columnOffset] ?? "";
          }

          return {
            ...current,
            [stickyTableHeaderCacheKey]: nextValues,
          };
        });
      })
      .catch((error) => {
        if (!isCancelled) {
          pushErrorToast(error);
        }
      });

    return () => {
      isCancelled = true;
    };
  }, [pushErrorToast, stickyTableHeaderCacheKey, stickyTableHeaderRequest, tableHeaderCache]);

  useEffect(() => {
    if (!tableRowHintTarget || !tableRowHintTargetKey) {
      pendingTableRowHintRequestIdRef.current += 1;
      setTableRowHint(null);
      formulaRowContextScrollLeftRef.current = 0;
      return;
    }

    const requestId = pendingTableRowHintRequestIdRef.current + 1;
    const requests = buildTableRowHintRangeRequests(tableRowHintTarget);
    const tableKey = getTableRowHintTableKey(tableRowHintTarget);

    pendingTableRowHintRequestIdRef.current = requestId;

    setTableRowHint((current) => {
      if (current?.tableKey === tableKey) {
        return current;
      }

      formulaRowContextScrollLeftRef.current = 0;
      return null;
    });

    void Promise.all([
      window.appShell.getSheetDisplayRange(requests.header),
      window.appShell.getSheetDisplayRange(requests.row),
    ])
      .then(([headerRange, rowRange]) => {
        if (pendingTableRowHintRequestIdRef.current !== requestId) {
          return;
        }

        setTableRowHint({
          items: buildTableRowHintItems(
            tableRowHintTarget,
            headerRange.values[0] ?? [],
            rowRange.values[0] ?? [],
          ),
          tableKey,
          tableName: tableRowHintTarget.table.name,
          tableRowNumber: tableRowHintTarget.tableRowNumber,
        });
      })
      .catch((error) => {
        if (pendingTableRowHintRequestIdRef.current === requestId) {
          pushErrorToast(error);
        }
      });
  }, [pushErrorToast, tableRowHintTarget, tableRowHintTargetKey]);

  const applyTransaction = useCallback(
    async (operations: Parameters<typeof window.appShell.applyTransaction>[0]["operations"]) => {
      const result = await window.appShell.applyTransaction({ operations });

      setSheetSummary(result.summary);

      return result;
    },
    [],
  );

  const revealSearchResult = useCallback((result: WorkbookSearchResult) => {
    const cell: Item = [result.columnIndex, result.rowIndex];

    setGridSelection(createCellSelection(cell));
    requestAnimationFrame(() => {
      gridRef.current?.scrollTo(result.columnIndex, result.rowIndex, "both", 2, 4, {
        hAlign: "center",
        vAlign: "center",
      });
      gridRef.current?.focus();
    });
  }, []);

  const focusSearchResult = useCallback(
    (result: WorkbookSearchResult | null) => {
      if (!result) {
        return;
      }

      pendingSearchResultRef.current = result;

      if (activeSheet?.id === result.sheetId) {
        pendingSearchResultRef.current = null;
        revealSearchResult(result);
        return;
      }

      void applyTransaction([
        {
          sheetId: result.sheetId,
          type: "setActiveSheet",
        },
      ]).catch((error) => {
        pendingSearchResultRef.current = null;
        pushErrorToast(error);
      });
    },
    [activeSheet?.id, applyTransaction, pushErrorToast, revealSearchResult],
  );

  const updateSearchState = useCallback(
    (nextSearchState: WorkbookSearchState, shouldFocusResult = true) => {
      setSearchState(nextSearchState);
      setViewNonce((current) => current + 1);

      if (shouldFocusResult) {
        focusSearchResult(nextSearchState.activeResult);
      }
    },
    [focusSearchResult],
  );

  const openWorkbookSearch = useCallback(() => {
    setIsSearchOpen(true);

    void window.appShell
      .getSearchState()
      .then((nextSearchState) => {
        updateSearchState(nextSearchState, false);
      })
      .catch((error) => {
        pushErrorToast(error);
      });
  }, [pushErrorToast, updateSearchState]);

  const closeWorkbookSearch = useCallback(() => {
    setIsSearchOpen(false);

    void window.appShell
      .clearSearch()
      .then((nextSearchState) => {
        updateSearchState(nextSearchState, false);
      })
      .catch((error) => {
        pushErrorToast(error);
      });
  }, [pushErrorToast, updateSearchState]);

  const changeWorkbookSearchQuery = useCallback(
    (
      text: string,
      scope: WorkbookSearchScope,
      valueMode: WorkbookSearchValueMode,
      caseSensitive: boolean,
      wholeWord: boolean,
    ) => {
      void window.appShell
        .setSearchQuery({
          caseSensitive,
          scope,
          text,
          valueMode,
          wholeWord,
        })
        .then((nextSearchState) => {
          updateSearchState(nextSearchState, false);
        })
        .catch((error) => {
          pushErrorToast(error);
        });
    },
    [pushErrorToast, updateSearchState],
  );

  const goToSearchResult = useCallback(
    (direction: "next" | "previous") => {
      const request =
        direction === "next"
          ? window.appShell.goToNextSearchResult()
          : window.appShell.goToPreviousSearchResult();

      void request.then(updateSearchState).catch((error) => {
        pushErrorToast(error);
      });
    },
    [pushErrorToast, updateSearchState],
  );

  const handleColumnResize = useCallback(
    (_column: GridColumn, newSize: number, columnIndex: number) => {
      setColumnResizeOverrides((current) => ({
        ...current,
        [String(columnIndex)]: Math.round(newSize),
      }));
    },
    [],
  );

  const handleColumnResizeEnd = useCallback(
    (_column: GridColumn, newSize: number, columnIndex: number) => {
      if (!activeSheet) {
        return;
      }

      const width = Math.round(newSize);

      setColumnResizeOverrides((current) => ({
        ...current,
        [String(columnIndex)]: width,
      }));

      void applyTransaction([
        {
          columnIndex,
          sheetId: activeSheet.id,
          type: "setColumnWidth",
          width,
        },
      ])
        .then(() => {
          setColumnResizeOverrides((current) =>
            removeColumnResizeOverride(current, columnIndex, width),
          );
        })
        .catch((error) => {
          setColumnResizeOverrides((current) =>
            removeColumnResizeOverride(current, columnIndex, width),
          );
          pushErrorToast(error);
        });
    },
    [activeSheet, applyTransaction, pushErrorToast],
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
        if (!didLogInitialRangeStart) {
          didLogInitialRangeStart = true;
          window.appShell.logStartupTiming(
            "initial-range-request-start",
            `sheetId=${request.sheetId} startColumn=${request.startColumn} startRow=${request.startRow} columnCount=${request.columnCount} rowCount=${request.rowCount}`,
          );
        }

        const [rawRange, displayRange, styleRange] = await Promise.all([
          window.appShell.getSheetRange(request),
          window.appShell.getSheetDisplayRange(request),
          window.appShell.getSheetStyleRange(request),
        ]);

        if (pendingRangeRequestIdRef.current !== requestId) {
          return;
        }

        rawRangeCacheRef.current = rawRange;
        displayRangeCacheRef.current = displayRange;
        styleRangeCacheRef.current = styleRange;
        setViewNonce((current) => current + 1);

        if (!didLogInitialRangeDone) {
          didLogInitialRangeDone = true;
          window.appShell.logStartupTiming(
            "initial-range-request-done",
            `sheetId=${request.sheetId} columnCount=${request.columnCount} rowCount=${request.rowCount}`,
          );
        }
      } catch (error) {
        if (!didLogInitialRangeDone) {
          didLogInitialRangeDone = true;
          window.appShell.logStartupTiming(
            "initial-range-request-failed",
            error instanceof Error ? error.message : "unknown error",
          );
        }

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

    if (!isCellWithinSheetBounds(selectedCell, activeSheet)) {
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
      const style = getCachedCellStyle(
        styleRangeCacheRef.current,
        columnIndex,
        rowIndex,
        activeSheet?.id,
      );

      if (rawValue === undefined || displayValue === undefined) {
        return createLoadingCell();
      }

      const isActiveSearchResult =
        activeSearchResultKey === `${activeSheet?.id}:${rowIndex}:${columnIndex}`;
      const tableHighlightStyle = getTableColumnHighlightStyle(
        activeSheetTableEntries,
        columnIndex,
        rowIndex,
        displayValue,
      );
      const themeOverride = mergeCellThemeOverrides(
        getCellThemeOverride(tableHighlightStyle),
        isActiveSearchResult ? SEARCH_ACTIVE_CELL_THEME : undefined,
      );

      return createTextCell(rawValue, displayValue, style, themeOverride);
    },
    [activeSearchResultKey, activeSheet?.id, activeSheetTableEntries, viewNonce],
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
        const [rawRange, displayRange, styleRange] = await Promise.all([
          window.appShell.getSheetRange(request),
          window.appShell.getSheetDisplayRange(request),
          window.appShell.getSheetStyleRange(request),
        ]);

        return displayRange.values.map((row, rowOffset) =>
          row.map((displayValue, columnOffset) => {
            const columnIndex = selection.x + columnOffset;
            const rowIndex = selection.y + rowOffset;
            const tableHighlightStyle = getTableColumnHighlightStyle(
              activeSheetTableEntries,
              columnIndex,
              rowIndex,
              displayValue,
            );

            return createTextCell(
              rawRange.values[rowOffset]?.[columnOffset] ?? displayValue,
              displayValue,
              styleRange.styles[rowOffset]?.[columnOffset] ?? undefined,
              getCellThemeOverride(tableHighlightStyle),
            );
          }),
        );
      };
    },
    [activeSheet, activeSheetTableEntries],
  );

  const provideGridEditor = useCallback<NonNullable<DataEditorProps["provideEditor"]>>((cell) => {
    if (cell.kind !== GridCellKind.Text) {
      return undefined;
    }

    return {
      disablePadding: cell.allowWrapping === true,
      editor: (props) => {
        const value = props.value as TextCell;

        return (
          <TextCellEntry
            altNewline
            autoFocus={value.readonly !== true}
            disabled={value.readonly === true}
            highlight={props.isHighlighted}
            onChange={(event) => {
              props.onChange({
                ...value,
                data: event.target.value,
              });
            }}
            onKeyDown={(event) => {
              const movement = getCellEditArrowKeyMovement(event);

              if (!movement) {
                return;
              }

              event.preventDefault();
              event.stopPropagation();
              props.onFinishedEditing(
                {
                  ...value,
                  data: event.currentTarget.value,
                },
                movement,
              );
            }}
            style={value.allowWrapping === true ? { padding: "3px 8.5px" } : undefined}
            validatedSelection={props.validatedSelection}
            value={value.data}
          />
        );
      },
    };
  }, []);

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
    (
      target: Item,
      values: readonly (readonly string[])[],
      options: { clipboard?: ClipboardRangePayload; mode?: ClipboardRangeMode } = {},
    ) => {
      if (!activeSheet || values.length === 0) {
        return false;
      }

      const [startColumn, startRow] = target;
      const nextValues = values.map((row) => [...row]);

      for (let rowOffset = 0; rowOffset < nextValues.length; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < nextValues[rowOffset].length; columnOffset += 1) {
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

      const selectedPastedValue = getPastedCellValue(target, selectedCell, values);

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

      void window.appShell
        .pasteRange({
          clipboard: options.clipboard,
          mode: options.mode,
          sheetId: activeSheet.id,
          startColumn,
          startRow,
          values: nextValues,
        })
        .then((result) => {
          setSheetSummary(result.summary);
        })
        .catch((error) => {
          pushErrorToast(error);
          void loadVisibleRange(lastVisibleRegionRef.current);
          void refreshSelectedCellData();
        });

      return true;
    },
    [activeSheet, loadVisibleRange, pushErrorToast, refreshSelectedCellData, selectedCell],
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
      input.setSelectionRange(nextState.selectionStart, nextState.selectionStart);
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

    if (selectionStart === selectionEnd && selectionStart >= input.value.length) {
      return false;
    }

    const deleteEnd = selectionStart === selectionEnd ? selectionStart + 1 : selectionEnd;
    const value = input.value.slice(0, selectionStart) + input.value.slice(deleteEnd);

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
        const result = await window.appShell.copyRange({
          ...currentSelectionRange,
          mode,
        });

        await window.appShell.writeClipboard({
          payload: result.clipboard,
          text: result.text,
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
        const result = await window.appShell.cutRange({
          ...currentSelectionRange,
          mode,
        });

        await window.appShell.writeClipboard({
          payload: result.clipboard,
          text: result.text,
        });

        setSheetSummary(result.summary);

        if (selectedCell && selectionContainsCell(gridSelection, selectedCell)) {
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

        if (nextText.length === 0 && !clipboard.payload && clipboard.text.length === 0) {
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

      return handlePaste(selectedCell, values, {
        clipboard: clipboard.payload,
        mode,
      });
    },
    [handlePaste, replaceFormulaInputSelection, selectedCell],
  );

  const restoreWorkbookHistory = useCallback(
    async (direction: "redo" | "undo") => {
      if (document.activeElement === formulaInputRef.current) {
        document.execCommand(direction);
        return true;
      }

      try {
        const result =
          direction === "undo" ? await window.appShell.undo() : await window.appShell.redo();

        setSheetSummary(result.summary);
        setSelectedChartId(null);
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
        return true;
      } catch (error) {
        pushErrorToast(error);
        return false;
      }
    },
    [loadVisibleRange, pushErrorToast, refreshSelectedCellData],
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

      const contextTable = getTableContainingCell(activeSheetTableEntries, cell);

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
          canFormat: true,
          canDeleteTable: Boolean(contextTable),
          canInsertTable: activeSheet !== null && !contextTable,
          canSortTable: Boolean(contextTable),
        })
        .catch((error) => {
          pushErrorToast(error);
        });
    },
    [activeSheet, activeSheetTableEntries, gridSelection, pushErrorToast],
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

  const openCellFormatDialog = useCallback(() => {
    if (!activeSheet || !sheetSummary || isModalDialogOpen) {
      return;
    }

    const ranges = getSelectedStyleRanges(
      gridSelection,
      activeSheet.id,
      activeSheet.rowCount,
      activeSheet.columnCount,
    );

    if (ranges.length === 0) {
      return;
    }

    setCellFormatSession({
      expectedVersion: sheetSummary.version,
      initialStyle: selectedCellData?.style,
      ranges,
    });
  }, [activeSheet, gridSelection, isModalDialogOpen, selectedCellData?.style, sheetSummary]);

  const clearSelectionFormatting = useCallback(() => {
    if (!activeSheet) {
      return false;
    }

    const ranges = getSelectedStyleRanges(
      gridSelection,
      activeSheet.id,
      activeSheet.rowCount,
      activeSheet.columnCount,
    );

    if (ranges.length === 0) {
      return false;
    }

    setSelectedCellData((current) =>
      current
        ? {
            ...current,
            style: undefined,
          }
        : current,
    );

    void applyTransaction(
      ranges.map((range) => ({
        ...range,
        type: "clearRangeStyle",
      })),
    ).catch((error) => {
      pushErrorToast(error);
      void loadVisibleRange(lastVisibleRegionRef.current);
      void refreshSelectedCellData();
    });

    return true;
  }, [
    activeSheet,
    applyTransaction,
    gridSelection,
    loadVisibleRange,
    pushErrorToast,
    refreshSelectedCellData,
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

  const openRenameSheetDialog = useCallback(() => {
    if (!activeSheet || !sheetSummary || isModalDialogOpen) {
      return;
    }

    setRenameSheetSession({
      expectedVersion: sheetSummary.version,
      name: activeSheet.name,
      sheetId: activeSheet.id,
    });
  }, [activeSheet, isModalDialogOpen, sheetSummary]);

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

  const openSheetQuickOpen = useCallback(() => {
    if (!sheetSummary || sheetSummary.sheets.length === 0 || isModalDialogOpen) {
      return;
    }

    setIsSheetQuickOpenOpen(true);
  }, [isModalDialogOpen, sheetSummary]);

  const closeSheetQuickOpen = useCallback(() => {
    setIsSheetQuickOpenOpen(false);
  }, []);

  const selectSheetFromQuickOpen = useCallback(
    (sheetId: string) => {
      setIsSheetQuickOpenOpen(false);
      void applyTransaction([
        {
          sheetId,
          type: "setActiveSheet",
        },
      ]).catch((error) => {
        pushErrorToast(error);
      });
    },
    [applyTransaction, pushErrorToast],
  );

  const handleFormulaInputChange = useCallback((event: ChangeEvent<HTMLInputElement>) => {
    setFormulaInputValue(event.target.value);
  }, []);

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

      const result = await window.appShell.saveWorkbookFile(sheetSummary.documentFilePath);

      setSheetSummary(result.summary);
    } catch (error) {
      pushErrorToast(error);
    }
  }, [handleSaveWorkbookAs, pushErrorToast, sheetSummary]);

  const closeChartEditor = useCallback(() => {
    setChartEditorSession(null);
  }, []);

  const closeCellFormatDialog = useCallback(() => {
    setCellFormatSession(null);
  }, []);

  const closeInstallationDialog = useCallback(() => {
    setInstallationDialogMode(null);
  }, []);

  const closeAboutDialog = useCallback(() => {
    setIsAboutDialogOpen(false);
  }, []);

  const closeRenameSheetDialog = useCallback(() => {
    setRenameSheetSession(null);
  }, []);

  const handleRenameSheetSaved = useCallback((summary: WorkbookSummary) => {
    setSheetSummary(summary);
  }, []);

  const handleChartEditorVersionConflict = useCallback(
    (message: string) => {
      setChartEditorSession(null);
      setCellFormatSession(null);
      setInstallationDialogMode(null);
      setRenameSheetSession(null);
      pushErrorToast(new Error(message));
    },
    [pushErrorToast],
  );

  const openCreateChartEditor = useCallback(() => {
    if (!activeSheet || !sheetSummary || isModalDialogOpen) {
      return;
    }

    setChartEditorSession({
      expectedVersion: sheetSummary.version,
      request: {
        mode: "create",
        sheetId: activeSheet.id,
        sourceRange: currentSelectionRange
          ? {
              columnCount: currentSelectionRange.columnCount,
              rowCount: currentSelectionRange.rowCount,
              sheetId: activeSheet.id,
              startColumn: currentSelectionRange.startColumn,
              startRow: currentSelectionRange.startRow,
            }
          : undefined,
      },
    });
  }, [activeSheet, currentSelectionRange, isModalDialogOpen, sheetSummary]);

  const createTableFromSelection = useCallback(() => {
    if (!activeSheet || !currentSelectionRange) {
      return false;
    }

    void applyTransaction([
      {
        hasHeaderRow: true,
        range: {
          columnCount: currentSelectionRange.columnCount,
          rowCount: currentSelectionRange.rowCount,
          sheetId: activeSheet.id,
          startColumn: currentSelectionRange.startColumn,
          startRow: currentSelectionRange.startRow,
        },
        type: "addTable",
      },
    ]).catch((error) => {
      pushErrorToast(error);
    });

    return true;
  }, [activeSheet, applyTransaction, currentSelectionRange, pushErrorToast]);

  const deleteSelectedTable = useCallback(() => {
    if (!selectedTable) {
      return false;
    }

    void applyTransaction([
      {
        tableId: selectedTable.id,
        type: "deleteTable",
      },
    ])
      .then(() => {
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
      })
      .catch((error) => {
        pushErrorToast(error);
        void loadVisibleRange(lastVisibleRegionRef.current);
        void refreshSelectedCellData();
      });

    return true;
  }, [applyTransaction, loadVisibleRange, pushErrorToast, refreshSelectedCellData, selectedTable]);

  const sortTableColumn = useCallback(
    (tableId: string, columnIndex: number, direction: "ascending" | "descending") => {
      void applyTransaction([createSortTableOperation(tableId, columnIndex, direction)])
        .then(() => {
          void loadVisibleRange(lastVisibleRegionRef.current);
          void refreshSelectedCellData();
        })
        .catch((error) => {
          pushErrorToast(error);
          void loadVisibleRange(lastVisibleRegionRef.current);
          void refreshSelectedCellData();
        });

      return true;
    },
    [applyTransaction, loadVisibleRange, pushErrorToast, refreshSelectedCellData],
  );

  const sortSelectedTable = useCallback(
    (direction: "ascending" | "descending") => {
      if (!selectedTable || !selectedCell) {
        return false;
      }

      return sortTableColumn(selectedTable.id, selectedCell[0], direction);
    },
    [selectedCell, selectedTable, sortTableColumn],
  );

  const openEditChartEditor = useCallback(
    (chartId: string) => {
      if (!sheetSummary || isModalDialogOpen) {
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
    [isModalDialogOpen, sheetSummary],
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
      const nextLayout = {
        ...layout,
        zIndex: Math.max(layout.zIndex, currentMaxZIndex + 1),
      };
      const previousSheetChartPreviews = sheetChartPreviews;

      setSheetChartPreviews((current) =>
        current
          ? {
              ...current,
              previews: current.previews.map((preview) =>
                preview.chart.id === chartId
                  ? {
                      ...preview,
                      chart: {
                        ...preview.chart,
                        layout: nextLayout,
                      },
                    }
                  : preview,
              ),
            }
          : current,
      );

      try {
        const result = await window.appShell.applyTransaction({
          expectedVersion: sheetSummary.version,
          operations: [
            {
              chartId,
              layout: nextLayout,
              type: "setChartLayout",
            },
          ],
        });

        setSheetSummary(result.summary);
      } catch (error) {
        setSheetChartPreviews((current) => {
          if (!current || !previousSheetChartPreviews) {
            return current;
          }

          const currentPreview = current.previews.find((preview) => preview.chart.id === chartId);

          if (currentPreview && !chartLayoutsEqual(currentPreview.chart.layout, nextLayout)) {
            return current;
          }

          return previousSheetChartPreviews;
        });
        pushErrorToast(error);
      }
    },
    [pushErrorToast, sheetChartPreviews, sheetSummary],
  );

  useEffect(() => {
    let isMounted = true;

    if (!didLogInitialWorkbookSummaryStart) {
      didLogInitialWorkbookSummaryStart = true;
      window.appShell.logStartupTiming("initial-workbook-summary-request-start");
    }

    void window.appShell
      .getWorkbookSummary()
      .then((summary) => {
        if (!isMounted) {
          return;
        }

        setSheetSummary(summary);

        if (!didLogInitialWorkbookSummaryDone) {
          didLogInitialWorkbookSummaryDone = true;
          window.appShell.logStartupTiming(
            "initial-workbook-summary-request-done",
            `version=${summary.version} sheets=${summary.sheets.length} charts=${summary.charts.length}`,
          );
        }
      })
      .catch((error) => {
        if (!isMounted) {
          return;
        }

        if (!didLogInitialWorkbookSummaryDone) {
          didLogInitialWorkbookSummaryDone = true;
          window.appShell.logStartupTiming(
            "initial-workbook-summary-request-failed",
            error instanceof Error ? error.message : "unknown error",
          );
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

    void window.appShell.setChartDialogOpen(isModalDialogOpen).catch((error) => {
      if (!isCancelled) {
        pushErrorToast(error);
      }
    });

    return () => {
      isCancelled = true;

      if (isModalDialogOpen) {
        void window.appShell.setChartDialogOpen(false);
      }
    };
  }, [isModalDialogOpen, pushErrorToast]);

  useEffect(() => {
    setColumnResizeOverrides({});
    setGridSelection(createEmptyGridSelection());
    setSelectedCellData(null);
    setFormulaInputValue("");
  }, [activeSheet?.id]);

  useEffect(() => {
    const pendingResult = pendingSearchResultRef.current;

    if (!pendingResult || pendingResult.sheetId !== activeSheet?.id) {
      return;
    }

    pendingSearchResultRef.current = null;
    revealSearchResult(pendingResult);
  }, [activeSheet?.id, revealSearchResult]);

  useEffect(() => {
    if (!isSearchOpen || searchState.query.text.length === 0) {
      return;
    }

    let isCancelled = false;

    void window.appShell
      .getSearchState()
      .then((nextSearchState) => {
        if (!isCancelled) {
          updateSearchState(nextSearchState, false);
        }
      })
      .catch((error) => {
        if (!isCancelled) {
          pushErrorToast(error);
        }
      });

    return () => {
      isCancelled = true;
    };
  }, [
    isSearchOpen,
    pushErrorToast,
    searchState.query.text.length,
    sheetSummary?.activeSheetId,
    sheetSummary?.version,
    updateSearchState,
  ]);

  useEffect(() => {
    if (!selectedCell || !activeSheet) {
      setSelectedCellData(null);
      setFormulaInputValue("");
      return;
    }

    if (!isCellWithinSheetBounds(selectedCell, activeSheet)) {
      setSelectedCellData(null);
      setFormulaInputValue("");
      return;
    }

    void refreshSelectedCellData();
  }, [activeSheet, refreshSelectedCellData, selectedCell, sheetSummary?.version]);

  useEffect(() => {
    if (!activeSheet) {
      setSheetChartPreviews(null);
      setSelectedChartId(null);
      setIsSheetChartPreviewsLoading(false);
      return;
    }

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

    if (!didLogInitialChartPreviewsStart) {
      didLogInitialChartPreviewsStart = true;
      window.appShell.logStartupTiming("initial-chart-previews-request-start", activeSheet.id);
    }

    void window.appShell
      .getSheetChartPreviews(activeSheet.id)
      .then((result) => {
        if (pendingSheetChartPreviewsRequestIdRef.current !== requestId) {
          return;
        }

        setSheetChartPreviews(result);
        setSelectedChartId((current) =>
          result.previews.some((preview) => preview.chart.id === current) ? current : null,
        );

        if (!didLogInitialChartPreviewsDone) {
          didLogInitialChartPreviewsDone = true;
          window.appShell.logStartupTiming(
            "initial-chart-previews-request-done",
            `sheetId=${activeSheet.id} charts=${result.previews.length}`,
          );
        }
      })
      .catch((error) => {
        if (pendingSheetChartPreviewsRequestIdRef.current !== requestId) {
          return;
        }

        setSheetChartPreviews(null);
        setSelectedChartId(null);

        if (!didLogInitialChartPreviewsDone) {
          didLogInitialChartPreviewsDone = true;
          window.appShell.logStartupTiming(
            "initial-chart-previews-request-failed",
            error instanceof Error ? error.message : "unknown error",
          );
        }

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
      if (isModalDialogOpen) {
        return;
      }

      const handleMenuAction = (nextAction: AppMenuAction) => {
        switch (nextAction) {
          case APP_MENU_ACTIONS.about:
            setIsAboutDialogOpen(true);
            return;
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
          case APP_MENU_ACTIONS.undo:
            void restoreWorkbookHistory("undo");
            return;
          case APP_MENU_ACTIONS.redo:
            void restoreWorkbookHistory("redo");
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
          case APP_MENU_ACTIONS.installation:
            setInstallationDialogMode("manage");
            return;
          case APP_MENU_ACTIONS.checkUpdates:
            setInstallationDialogMode("check-updates");
            return;
          case APP_MENU_ACTIONS.paste:
            void pasteSelection("raw");
            return;
          case APP_MENU_ACTIONS.pasteValues:
            void pasteSelection("display");
            return;
          case APP_MENU_ACTIONS.formatCells:
            openCellFormatDialog();
            return;
          case APP_MENU_ACTIONS.clearFormatting:
            clearSelectionFormatting();
            return;
          case APP_MENU_ACTIONS.find:
            openWorkbookSearch();
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
          case APP_MENU_ACTIONS.renameSheet:
            openRenameSheetDialog();
            return;
          case APP_MENU_ACTIONS.deleteSheet:
            deleteSheet();
            return;
          case APP_MENU_ACTIONS.selectSheet:
            openSheetQuickOpen();
            return;
          case APP_MENU_ACTIONS.insertChart:
            openCreateChartEditor();
            return;
          case APP_MENU_ACTIONS.insertTable:
            createTableFromSelection();
            return;
          case APP_MENU_ACTIONS.deleteTable:
            deleteSelectedTable();
            return;
          case APP_MENU_ACTIONS.sortTableAscending:
            sortSelectedTable("ascending");
            return;
          case APP_MENU_ACTIONS.sortTableDescending:
            sortSelectedTable("descending");
            return;
        }
      };

      handleMenuAction(action);
    });
  }, [
    addColumn,
    addRow,
    addSheet,
    clearSelectionFormatting,
    cutSelection,
    copySelection,
    createTableFromSelection,
    deleteSelection,
    deleteSelectedTable,
    deleteSheet,
    handleExport,
    handleImport,
    handleOpenWorkbook,
    handleSaveWorkbook,
    handleSaveWorkbookAs,
    isModalDialogOpen,
    openCellFormatDialog,
    openCreateChartEditor,
    openSheetQuickOpen,
    openRenameSheetDialog,
    openWorkbookSearch,
    pasteSelection,
    restoreWorkbookHistory,
    sortSelectedTable,
  ]);

  useEffect(() => {
    const handleWindowKeyDown = (event: globalThis.KeyboardEvent) => {
      if (isModalDialogOpen) {
        return;
      }

      const isPrimaryModifier = event.ctrlKey || event.metaKey;
      const activeElement = document.activeElement;
      const isFormulaInputFocused = activeElement === formulaInputRef.current;
      const normalizedKey = event.key.toLowerCase();

      if (event.altKey) {
        return;
      }

      if (
        shouldCloseWorkbookSearchOnEscape(event, {
          hasSearchQuery: searchState.query.text.length > 0,
          isFormulaInputFocused,
          isSearchOpen,
        })
      ) {
        event.preventDefault();
        closeWorkbookSearch();
        return;
      }

      if (isPrimaryModifier && !event.shiftKey && normalizedKey === "p") {
        event.preventDefault();
        openSheetQuickOpen();
        return;
      }

      if (isPrimaryModifier && !event.shiftKey && normalizedKey === "f") {
        event.preventDefault();
        openWorkbookSearch();
        return;
      }

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

      if (
        isPrimaryModifier &&
        !isFormulaInputFocused &&
        (normalizedKey === "z" || normalizedKey === "y")
      ) {
        event.preventDefault();
        void restoreWorkbookHistory(event.shiftKey || normalizedKey === "y" ? "redo" : "undo");
        return;
      }

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

      if (event.key === "Delete" && !event.shiftKey && !event.ctrlKey && !event.metaKey) {
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
    closeWorkbookSearch,
    cutSelection,
    deleteSelection,
    isModalDialogOpen,
    isSearchOpen,
    openSheetQuickOpen,
    openWorkbookSearch,
    pasteSelection,
    restoreWorkbookHistory,
    searchState.query.text.length,
  ]);

  return (
    <main className="app-shell">
      <section className="formula-bar" aria-label="Formula bar">
        <div className="formula-bar__address">{selectedCellAddress || "Cell"}</div>
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
                ? "Type a value or a formula like =A1+'Sheet 2'!B1"
                : "Select a cell to inspect or edit"
            }
            value={selectedCell ? formulaInputValue : ""}
          />
        </div>
        {tableRowHint ? (
          <section
            aria-label={`Table row context for ${tableRowHint.tableName}`}
            className="formula-row-context"
          >
            <div className="formula-row-context__summary">
              <strong>{tableRowHint.tableName}</strong>
              <span>{`Row ${tableRowHint.tableRowNumber}`}</span>
            </div>
            <dl
              className="formula-row-context__items"
              onScroll={(event) => {
                formulaRowContextScrollLeftRef.current = event.currentTarget.scrollLeft;
              }}
              ref={formulaRowContextItemsRef}
            >
              {tableRowHint.items.map((item) => (
                <div
                  className={`formula-row-context__item${item.isActive ? " is-active" : ""}`}
                  key={item.columnIndex}
                >
                  <dt>{item.label}</dt>
                  <dd>
                    {item.value ? (
                      item.value
                    ) : (
                      <span className="formula-row-context__blank">Blank</span>
                    )}
                  </dd>
                </div>
              ))}
            </dl>
          </section>
        ) : null}
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
            highlightRegions={searchHighlightRegions}
            onCellEdited={handleCellEdited}
            maxColumnWidth={MAX_COLUMN_WIDTH}
            minColumnWidth={MIN_COLUMN_WIDTH}
            onColumnResize={handleColumnResize}
            onColumnResizeEnd={handleColumnResizeEnd}
            onDelete={(selection) => {
              deleteSelection(selection);
              return false;
            }}
            onGridSelectionChange={setGridSelection}
            onPaste={handlePaste}
            provideEditor={provideGridEditor}
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
          {isSearchOpen ? (
            <WorkbookSearchBox
              onChangeQuery={changeWorkbookSearchQuery}
              onClose={closeWorkbookSearch}
              onNext={() => {
                goToSearchResult("next");
              }}
              onPrevious={() => {
                goToSearchResult("previous");
              }}
              state={searchState}
            />
          ) : null}
          {tableSortControls.length > 0 ? (
            <div className="table-sort-overlay" aria-label="Table sort controls">
              {tableSortControls.map((control) => {
                const iconDirection = control.direction ?? "ascending";
                const nextDirection = getNextTableSortDirection(control.direction);

                return (
                  <div
                    className="table-sort-control"
                    key={`${control.tableId}:${control.columnIndex}`}
                    style={{
                      height: control.height,
                      left: control.left,
                      top: control.top,
                      width: control.width,
                    }}
                    onPointerDown={(event) => {
                      event.stopPropagation();
                    }}
                  >
                    <button
                      aria-label={`Sort ${getColumnTitle(control.columnIndex)} ${nextDirection}`}
                      className={`table-sort-control__button${control.direction ? " is-active" : ""}`}
                      title={`Sort ${nextDirection}`}
                      type="button"
                      onClick={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                        sortTableColumn(control.tableId, control.columnIndex, nextDirection);
                      }}
                    >
                      <span
                        className={`table-sort-control__icon table-sort-control__icon--${iconDirection}`}
                      />
                    </button>
                  </div>
                );
              })}
            </div>
          ) : null}
          {shouldRenderChartOverlay ? (
            <Suspense fallback={null}>
              <LazyWorkbookChartOverlay
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
            </Suspense>
          ) : null}
        </section>
      </div>

      {isSheetQuickOpenOpen ? (
        <SheetQuickOpen
          activeSheetId={sheetSummary?.activeSheetId ?? null}
          onClose={closeSheetQuickOpen}
          onSelect={selectSheetFromQuickOpen}
          sheets={sheetSummary?.sheets ?? []}
        />
      ) : null}

      {isAboutDialogOpen ? <AboutDialog onClose={closeAboutDialog} /> : null}

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

      {cellFormatSession ? (
        <CellFormatDialog
          expectedVersion={cellFormatSession.expectedVersion}
          initialStyle={cellFormatSession.initialStyle}
          key={`format:${cellFormatSession.expectedVersion}:${cellFormatSession.ranges
            .map(
              (range) =>
                `${range.sheetId}:${range.startRow}:${range.startColumn}:${range.rowCount}:${range.columnCount}`,
            )
            .join("|")}`}
          onClose={closeCellFormatDialog}
          onVersionConflict={handleChartEditorVersionConflict}
          ranges={cellFormatSession.ranges}
        />
      ) : null}

      {renameSheetSession ? (
        <RenameSheetDialog
          expectedVersion={renameSheetSession.expectedVersion}
          initialName={renameSheetSession.name}
          key={`rename:${renameSheetSession.sheetId}:${renameSheetSession.expectedVersion}`}
          onClose={closeRenameSheetDialog}
          onSaved={handleRenameSheetSaved}
          onVersionConflict={handleChartEditorVersionConflict}
          sheetId={renameSheetSession.sheetId}
        />
      ) : null}

      {installationDialogMode ? (
        <InstallationDialog
          initialMode={installationDialogMode}
          key={`installation:${installationDialogMode}`}
          onClose={closeInstallationDialog}
        />
      ) : null}

      <ToastViewport onDismiss={dismissToast} toasts={toasts} />
    </main>
  );
}
