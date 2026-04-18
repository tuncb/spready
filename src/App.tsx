import DataEditor, {
  GridCellKind,
  type EditableGridCell,
  type GridCell,
  type GridColumn,
  type Item,
} from '@glideapps/glide-data-grid';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

const INITIAL_ROWS = 200;
const INITIAL_COLUMNS = 50;
const DEFAULT_COLUMN_WIDTH = 140;

function createSheet(rowCount: number, columnCount: number): string[][] {
  return Array.from({ length: rowCount }, () => Array(columnCount).fill(''));
}

function normalizeSheet(rows: string[][]): string[][] {
  const rowCount = Math.max(rows.length, 1);
  const columnCount = Math.max(1, ...rows.map((row) => row.length));

  return Array.from({ length: rowCount }, (_, rowIndex) => {
    const sourceRow = rows[rowIndex] ?? [];

    return Array.from({ length: columnCount }, (_, columnIndex) => sourceRow[columnIndex] ?? '');
  });
}

function parseCsv(content: string): string[][] {
  if (content.length === 0) {
    return normalizeSheet([]);
  }

  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentValue = '';
  let isQuoted = false;

  for (let index = 0; index < content.length; index += 1) {
    const character = content[index];

    if (isQuoted) {
      if (character === '"') {
        if (content[index + 1] === '"') {
          currentValue += '"';
          index += 1;
        } else {
          isQuoted = false;
        }
      } else {
        currentValue += character;
      }

      continue;
    }

    if (character === '"') {
      isQuoted = true;
      continue;
    }

    if (character === ',') {
      currentRow.push(currentValue);
      currentValue = '';
      continue;
    }

    if (character === '\n' || character === '\r') {
      if (character === '\r' && content[index + 1] === '\n') {
        index += 1;
      }

      currentRow.push(currentValue);
      rows.push(currentRow);
      currentRow = [];
      currentValue = '';
      continue;
    }

    currentValue += character;
  }

  if (currentRow.length > 0 || currentValue.length > 0) {
    currentRow.push(currentValue);
    rows.push(currentRow);
  }

  return normalizeSheet(rows);
}

function getUsedSheetRange(sheet: string[][]): string[][] {
  let lastRowIndex = -1;
  let lastColumnIndex = -1;

  for (let rowIndex = 0; rowIndex < sheet.length; rowIndex += 1) {
    const row = sheet[rowIndex];

    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (row[columnIndex] === '') {
        continue;
      }

      lastRowIndex = rowIndex;
      lastColumnIndex = Math.max(lastColumnIndex, columnIndex);
    }
  }

  if (lastRowIndex < 0 || lastColumnIndex < 0) {
    return [];
  }

  return sheet
    .slice(0, lastRowIndex + 1)
    .map((row) => row.slice(0, lastColumnIndex + 1));
}

function escapeCsvValue(value: string): string {
  if (/[",\r\n]/.test(value)) {
    return `"${value.replaceAll('"', '""')}"`;
  }

  return value;
}

function serializeCsv(sheet: string[][]): string {
  const usedRange = getUsedSheetRange(sheet);

  return usedRange.map((row) => row.map(escapeCsvValue).join(',')).join('\r\n');
}

function getColumnTitle(index: number): string {
  let current = index;
  let label = '';

  do {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);

  return label;
}

function createColumns(columnCount: number): GridColumn[] {
  return Array.from({ length: columnCount }, (_, index) => ({
    id: `column-${index}`,
    title: getColumnTitle(index),
    width: DEFAULT_COLUMN_WIDTH,
  }));
}

function createTextCell(value: string): GridCell {
  return {
    kind: GridCellKind.Text,
    allowOverlay: true,
    data: value,
    displayData: value,
  };
}

function applyTextCellEdit(
  previousSheet: string[][],
  rowIndex: number,
  columnIndex: number,
  value: string,
): string[][] {
  const currentRow = previousSheet[rowIndex];

  if (!currentRow || columnIndex < 0 || columnIndex >= currentRow.length) {
    return previousSheet;
  }

  if (currentRow[columnIndex] === value) {
    return previousSheet;
  }

  const nextSheet = [...previousSheet];
  const nextRow = [...currentRow];

  nextRow[columnIndex] = value;
  nextSheet[rowIndex] = nextRow;

  return nextSheet;
}

export default function App() {
  const [sheet, setSheet] = useState(() => createSheet(INITIAL_ROWS, INITIAL_COLUMNS));
  const sheetRef = useRef(sheet);
  const filePathRef = useRef<string>();

  sheetRef.current = sheet;

  const rowCount = sheet.length;
  const columnCount = sheet[0]?.length ?? 0;
  const columns = useMemo(() => createColumns(columnCount), [columnCount]);

  const getCellContent = useCallback((cell: Item): GridCell => {
    const [columnIndex, rowIndex] = cell;
    const value = sheetRef.current[rowIndex]?.[columnIndex] ?? '';

    return createTextCell(value);
  }, []);

  const handleCellEdited = useCallback((cell: Item, newValue: EditableGridCell) => {
    if (newValue.kind !== GridCellKind.Text) {
      return;
    }

    const [columnIndex, rowIndex] = cell;

    setSheet((previousSheet) =>
      applyTextCellEdit(previousSheet, rowIndex, columnIndex, newValue.data),
    );
  }, []);

  const handlePaste = useCallback((target: Item, values: readonly (readonly string[])[]) => {
    if (values.length === 0) {
      return false;
    }

    const [startColumnIndex, startRowIndex] = target;

    setSheet((previousSheet) => {
      const maxColumnCount = previousSheet[0]?.length ?? 0;
      const nextSheet = [...previousSheet];
      let changed = false;

      for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
        const rowIndex = startRowIndex + rowOffset;

        if (rowIndex < 0 || rowIndex >= previousSheet.length) {
          continue;
        }

        const sourceRow = values[rowOffset];
        let nextRow = nextSheet[rowIndex];
        let rowCloned = false;

        for (let columnOffset = 0; columnOffset < sourceRow.length; columnOffset += 1) {
          const columnIndex = startColumnIndex + columnOffset;

          if (columnIndex < 0 || columnIndex >= maxColumnCount) {
            continue;
          }

          const pastedValue = sourceRow[columnOffset] ?? '';

          if (nextRow[columnIndex] === pastedValue) {
            continue;
          }

          if (!rowCloned) {
            nextRow = [...nextRow];
            nextSheet[rowIndex] = nextRow;
            rowCloned = true;
          }

          nextRow[columnIndex] = pastedValue;
          changed = true;
        }
      }

      return changed ? nextSheet : previousSheet;
    });

    return true;
  }, []);

  const addRow = useCallback(() => {
    setSheet((previousSheet) => [
      ...previousSheet,
      Array(previousSheet[0]?.length ?? 0).fill(''),
    ]);
  }, []);

  const addColumn = useCallback(() => {
    setSheet((previousSheet) => previousSheet.map((row) => [...row, '']));
  }, []);

  const handleImport = useCallback(async () => {
    const result = await window.appShell.openCsvFile();

    if (result.canceled) {
      return;
    }

    setSheet(parseCsv(result.content));
    filePathRef.current = result.filePath;
  }, []);

  const handleExport = useCallback(async () => {
    const defaultPath = filePathRef.current ?? 'Sheet1.csv';
    const result = await window.appShell.saveCsvFile(
      serializeCsv(sheetRef.current),
      defaultPath,
    );

    if (!result.canceled) {
      filePathRef.current = result.filePath;
    }
  }, []);

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
          <p className="app-shell__eyebrow">{window.appShell.name}</p>
          <h1 className="app-shell__title">Sheet 1</h1>
        </div>

        <div className="app-shell__stats" aria-label="Sheet size">
          <span>{rowCount} rows</span>
          <span>{columnCount} columns</span>
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
        </div>
      </header>

      <section className="sheet-surface" aria-label="Spreadsheet surface">
        <DataEditor
          columns={columns}
          getCellContent={getCellContent}
          getCellsForSelection
          height="100%"
          onCellEdited={handleCellEdited}
          onPaste={handlePaste}
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
