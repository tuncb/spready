# Spready

Spready now keeps workbook state in the Electron main process and exposes a local control endpoint for external LLM harnesses. The renderer is a view over that shared workbook state, not the source of truth.

Native workbook documents use the `.spready` extension. CSV import and export remain available as sheet-level interchange tools.

## Commands

### Install

```sh
npm install
```

### Run

```sh
npm start
```

When the app starts it also opens a local TCP control server on `127.0.0.1:45731` by default. If that port is busy it falls back to a random free local port and prints the chosen address in the Electron console. You can override the preferred port with `SPREADY_CONTROL_PORT`.

Workbook summaries report `hasUnsavedChanges` so clients can decide whether to save, discard, or cancel before replacing the current workbook.

### MCP stdio wrapper

Start the Electron app first, then run:

```sh
npm run mcp:stdio
```

The wrapper connects to the running app over the local control server and exposes MCP tools, resources, and prompts over stdio for external harnesses.

Connection discovery order:

- `--host` / `--port`
- `SPREADY_CONTROL_HOST` / `SPREADY_CONTROL_PORT`
- temp discovery file at `os.tmpdir()/spready-control.json`
- default `127.0.0.1:45731`

Example harness config:

```json
{
  "mcpServers": {
    "spready": {
      "command": "npm",
      "args": ["run", "mcp:stdio"]
    }
  }
}
```

Release bundles also include a standalone `spready-mcp` executable and a `spready.mcp.json`
template you can import into a harness after extracting the archive.

### Checks

```sh
npm run lint
npm run typecheck
```

### Package

```sh
npm run package
```

### Build distributables

```sh
npm run make
```

### Tests

```sh
npm test
```

## Control API

The control protocol is newline-delimited JSON over a local TCP socket.

Each request is a single JSON line:

```json
{ "id": 1, "method": "getWorkbookSummary" }
```

Each response is a single JSON line:

```json
{
  "id": 1,
  "ok": true,
  "result": {
    "activeSheetId": "sheet-1",
    "activeSheetName": "Sheet 1",
    "hasUnsavedChanges": false,
    "sheets": [
      { "id": "sheet-1", "name": "Sheet 1", "rowCount": 200, "columnCount": 50 }
    ],
    "version": 0
  }
}
```

On connect, the server sends a `hello` event. Workbook mutations also emit `workbookChanged` events to all connected clients.

### Methods

- `ping`
- `listMethods`
- `getControlInfo`
- `getWorkbookSummary`
- `getCellData`
- `getSheetDisplayRange`
- `getSheetRange`
- `getUsedRange`
- `getSheetCsv`
- `copyRange`
- `cutRange`
- `pasteRange`
- `clearRange`
- `createNewWorkbook`
- `openWorkbookFile`
- `saveWorkbookFile`
- `importCsvFile`
- `exportCsvFile`
- `applyTransaction`

For workbook-targeted methods, including `importCsvFile` and `exportCsvFile`, you can pass
`sheetId` explicitly. If you omit `sheetId`, the active sheet is used.

`copyRange` returns one rectangular range as tab-delimited text using raw input or displayed values. `cutRange` returns the same clipboard payloads and clears the source cells through the controller in one mutation.

### CSV import/export examples

Import a CSV file into a specific sheet without changing the active sheet:

```json
{
  "id": 3,
  "method": "importCsvFile",
  "params": {
    "filePath": "C:\\\\data\\\\quarterly.csv",
    "sheetId": "sheet-2"
  }
}
```

Export a specific sheet to CSV without changing the active sheet:

```json
{
  "id": 4,
  "method": "exportCsvFile",
  "params": {
    "filePath": "C:\\\\exports\\\\quarterly.csv",
    "sheetId": "sheet-2"
  }
}
```

### Workbook file examples

Create a new blank workbook, replacing the current workbook only if discarding unsaved changes is intentional:

```json
{
  "id": 4,
  "method": "createNewWorkbook",
  "params": {
    "discardUnsavedChanges": true
  }
}
```

Open a native Spready workbook file:

```json
{
  "id": 5,
  "method": "openWorkbookFile",
  "params": {
    "filePath": "C:\\\\workbooks\\\\budget.spready",
    "discardUnsavedChanges": true
  }
}
```

Save the current workbook as a native Spready workbook file:

```json
{
  "id": 6,
  "method": "saveWorkbookFile",
  "params": {
    "filePath": "C:\\\\workbooks\\\\budget.spready"
  }
}
```

### Transaction example

```json
{
  "id": 2,
  "method": "applyTransaction",
  "params": {
    "operations": [
      {
        "type": "setRange",
        "startRow": 0,
        "startColumn": 0,
        "values": [
          ["Name", "Revenue"],
          ["North", "1200"],
          ["South", "980"]
        ]
      },
      {
        "type": "insertRows",
        "rowIndex": 3,
        "count": 2
      }
    ]
  }
}
```

Supported transaction operations currently include:

- `addSheet`
- `addChart`
- `setActiveSheet`
- `setChartSpec`
- `renameSheet`
- `renameChart`
- `deleteSheet`
- `deleteChart`
- `resizeSheet`
- `insertRows`
- `deleteRows`
- `insertColumns`
- `deleteColumns`
- `setCell`
- `setRange`
- `clearRange`
- `replaceSheet`
- `replaceSheetFromCsv`
- `setSheetSourceFile`

## MCP surface

The stdio MCP wrapper currently exposes:

### Tools

- `describe_capabilities`
- `get_workbook_summary`
- `create_new_workbook`
- `get_used_range`
- `get_cell_data`
- `get_sheet_display_range`
- `get_sheet_range`
- `get_sheet_csv`
- `open_workbook_file`
- `save_workbook_file`
- `import_csv_file`
- `export_csv_file`
- `apply_transaction`

`get_sheet_range` returns raw stored cell input, including formula strings like `=A1+B1`.
`get_sheet_display_range` returns evaluated display values for the grid view.
`get_cell_data` returns both the raw input and the evaluated display value for one cell.

Display reads evaluate the same-sheet formula engine used by the app UI, including arithmetic, comparisons, text operators, ranges, core math/logical/text functions, and same-sheet lookup functions such as `INDEX`, `MATCH`, and `XLOOKUP`.
Raw reads continue to preserve the stored input exactly as written.

`import_csv_file` and `export_csv_file` both accept an optional `sheetId`. If omitted, they use
the active sheet.

`open_workbook_file` and `save_workbook_file` operate on the full multi-sheet workbook and use the `.spready` document format.

When `get_workbook_summary` reports `hasUnsavedChanges: true`, remote clients should either save first or pass `discardUnsavedChanges: true` to `open_workbook_file` only when replacing local changes is intended.

`create_new_workbook` follows the same rule and requires `discardUnsavedChanges: true` when replacing a dirty in-memory workbook.

### Resources

- `spready://guide`
- `spready://workbook/summary`

Clients that support MCP resource subscriptions can subscribe to `spready://workbook/summary`
and receive live `resources/updated` notifications when the workbook changes.

### Prompts

- `spready_workbook_task`
