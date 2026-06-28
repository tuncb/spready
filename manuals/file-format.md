# Spready File Format

Spready native workbook files use the `.spready` extension. A `.spready` file
is UTF-8 JSON, formatted with two-space indentation when saved by the app.

The current writer emits `formatVersion: 6`. The reader accepts versions `4`,
`5`, and `6`; new files should use version `6`.

## Top Level

```json
{
  "format": "spready-workbook",
  "formatVersion": 6,
  "workbook": {
    "activeSheetId": "sheet-1",
    "charts": [],
    "nextChartNumber": 1,
    "nextSheetNumber": 2,
    "nextTableNumber": 1,
    "sheets": [],
    "tables": []
  }
}
```

`format` must be exactly `spready-workbook`. `activeSheetId` must match one
sheet id. `nextSheetNumber`, `nextChartNumber`, and `nextTableNumber` are
positive integer counters for generated ids. `tables` and `nextTableNumber` may
be absent in older readable files; Spready restores them to safe defaults.

Runtime-only state such as the file path, undo history, dirty flag, evaluated
formula values, and workbook version is not stored.

## Coordinates

All row and column indexes in the file are zero-based. Counts are positive
except chart source ranges, which may have zero rows or columns for an invalid
but preserved chart.

## Sheets

Each workbook has at least one sheet:

```json
{
  "id": "sheet-1",
  "name": "Sheet 1",
  "rowCount": 200,
  "columnCount": 50,
  "cells": [{ "row": 0, "column": 0, "value": "Revenue" }],
  "styles": [],
  "columnWidths": { "1": 220 },
  "metadata": { "sourceFilePath": "C:\\imports\\data.csv" }
}
```

`id` and `name` are required non-empty strings. Sheet names must be unique
case-insensitively. `rowCount` and `columnCount` must be at least `1`.

Cells are sparse. Empty cells are omitted, and each listed cell must be inside
the sheet bounds with no duplicate `(row,column)` entry. `value` is the raw user
input string: formulas are stored literally, for example `"=A1+B1"`.

`columnWidths` is optional and sparse. Keys are zero-based column indexes as
strings; widths are pixels from `48` through `640`. Widths are rounded when
loaded. `metadata.sourceFilePath` is optional import provenance.

## Cell Styles

`styles` is a sparse array of styled cells:

```json
{
  "row": 0,
  "column": 1,
  "style": {
    "bold": true,
    "horizontalAlign": "right",
    "numberFormat": { "type": "percent", "decimalPlaces": 1 }
  }
}
```

Each style entry must be in bounds, unique per cell, and contain at least one
style property. Supported properties are `backgroundColor`, `bold`,
`fontFamily`, `fontSize`, `horizontalAlign`, `italic`, `numberFormat`,
`textColor`, and `wrapText`.

`horizontalAlign` is `left`, `center`, or `right`. `fontSize` must be at least
`6`. Number formats are:

- `{ "type": "general" }`
- `{ "type": "number", "decimalPlaces": 2, "useGrouping": true }`
- `{ "type": "percent", "significantDigits": 3 }`
- `{ "type": "scientific", "significantDigits": 4 }`

For `number` and `percent`, use either `decimalPlaces` (`0` to `20`) or
`significantDigits` (`1` to `21`), not both.

## Tables

Tables persist rectangular sheet ranges:

```json
{
  "id": "table-1",
  "name": "Revenue",
  "hasHeaderRow": true,
  "range": {
    "sheetId": "sheet-1",
    "startRow": 0,
    "startColumn": 0,
    "rowCount": 10,
    "columnCount": 3
  },
  "sortState": {
    "keys": [{ "columnIndex": 1, "direction": "descending" }],
    "valueMode": "display"
  },
  "columnHighlightRules": [{ "columnIndex": 2, "threshold": 1000, "thresholdType": "higher" }]
}
```

Table ids must be unique. Table names must be unique case-insensitively. Ranges
must fit inside their sheet and cannot overlap another table on the same sheet.
`sortState` is optional. Sort keys use absolute sheet column indexes that must
fall inside the table range; directions are `ascending` or `descending`, and
`valueMode` is `raw` or `display`. `columnHighlightRules` is optional. Highlight
rules use absolute sheet column indexes inside the table range and compare
numeric displayed values with `threshold`. `thresholdType` is optional and
defaults to `higher`; valid values are `lower`, `lowerOrEqual`, `equal`,
`higherOrEqual`, and `higher`. Multiple rules may target one column. Matching
rule styles merge in order, and later rules override earlier style fields.
Rules may include `backgroundColor`, `textColor`, and `bold`.

## Charts

Charts are embedded objects owned by a sheet:

```json
{
  "id": "chart-1",
  "name": "Revenue",
  "sheetId": "sheet-1",
  "layout": {
    "startRow": 0,
    "startColumn": 4,
    "offsetX": 0,
    "offsetY": 0,
    "width": 420,
    "height": 260,
    "zIndex": 0
  },
  "spec": {
    "family": "cartesian",
    "chartType": "bar",
    "source": {
      "range": {
        "sheetId": "sheet-1",
        "startRow": 0,
        "startColumn": 0,
        "rowCount": 10,
        "columnCount": 2
      },
      "seriesLayoutBy": "column",
      "sourceHeader": true
    },
    "categoryDimension": 0,
    "valueDimensions": [1]
  }
}
```

Chart ids must be unique. Layout `width` must be at least `180`, `height` at
least `140`, offsets cannot be negative, and `zIndex` cannot be negative.

Cartesian charts use `family: "cartesian"` and `chartType` of `bar`, `line`,
`area`, or `scatter`. They define `categoryDimension` and `valueDimensions`.
Optional flags are `smooth` and `stacked`.

Pie charts use:

```json
{
  "family": "pie",
  "chartType": "pie",
  "nameDimension": 0,
  "valueDimension": 1,
  "source": {
    "range": {
      "sheetId": "sheet-1",
      "startRow": 0,
      "startColumn": 0,
      "rowCount": 10,
      "columnCount": 2
    },
    "seriesLayoutBy": "column",
    "sourceHeader": true
  }
}
```

`seriesLayoutBy` is `column` or `row`. Chart references are preserved even if
their source or owner sheet is invalid; display/read APIs report validation
status separately.

## Validation Summary

Spready rejects files that are not valid JSON, have the wrong `format`, use an
unsupported `formatVersion`, contain duplicate sheet/chart/table ids, duplicate
case-insensitive sheet or table names, out-of-bounds cells/styles/column widths,
duplicate cell/style entries, missing active sheet, missing table sheet,
out-of-bounds table ranges, sort keys, or highlight columns, or overlapping
tables.

## Access Through Automation

Use `openWorkbookFile` / `saveWorkbookFile` over TCP or
`open_workbook_file` / `save_workbook_file` over MCP to read and write complete
`.spready` documents. Use `list_manuals` and `read_manual` over MCP to discover
and read this manual.
