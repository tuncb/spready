# Spready

Spready now keeps workbook state in the Electron main process and exposes a local control endpoint for external LLM harnesses. The renderer is a view over that shared workbook state, not the source of truth.

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

### MCP stdio wrapper

Start the Electron app first, then run:

```sh
npm run mcp:stdio
```

The wrapper connects to the running app over the local control server and exposes MCP tools over stdio for external harnesses.

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
# No automated test command is configured yet.
```

## Control API

The control protocol is newline-delimited JSON over a local TCP socket.

Each request is a single JSON line:

```json
{"id":1,"method":"getWorkbookSummary"}
```

Each response is a single JSON line:

```json
{"id":1,"ok":true,"result":{"activeSheetId":"sheet-1","activeSheetName":"Sheet 1","sheets":[{"id":"sheet-1","name":"Sheet 1","rowCount":200,"columnCount":50}],"version":0}}
```

On connect, the server sends a `hello` event. Workbook mutations also emit `workbookChanged` events to all connected clients.

### Methods

- `ping`
- `listMethods`
- `getControlInfo`
- `getWorkbookSummary`
- `getSheetRange`
- `getUsedRange`
- `getSheetCsv`
- `applyTransaction`

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
- `setActiveSheet`
- `renameSheet`
- `deleteSheet`
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

## MCP tools

The stdio MCP wrapper currently exposes:

- `get_workbook_summary`
- `get_used_range`
- `get_sheet_range`
- `get_sheet_csv`
- `apply_transaction`
