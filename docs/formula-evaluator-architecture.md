# Formula Evaluator Architecture

## Goal

Add a first-pass formula system that works the same way in all app surfaces:

- The workbook stores the raw cell input exactly as the user typed it.
- CSV import and export round-trip the raw formula string.
- The spreadsheet grid shows the evaluated result.
- The formula bar shows and edits the raw formula string.
- The same capability is available through the Electron UI, the TCP control server, and the MCP wrapper.

Initial scope:

- Formula prefix: `=`
- Operators: `+`, `-`, `*`, `/`
- Parentheses
- Same-sheet cell references such as `A1`, `B12`, `AA3`

Out of scope for v1:

- Functions like `SUM`
- Cross-sheet references
- Ranges like `A1:B5`
- Absolute references like `$A$1`
- Formula rewriting during row or column insert and delete

## Current Architecture

The current codebase already has a good split for this feature:

- [`src/workbook-core.ts`](/c:/work/spready/src/workbook-core.ts) is the source of truth for workbook structure, CSV parsing and serialization, and transaction application.
- [`src/workbook-controller.ts`](/c:/work/spready/src/workbook-controller.ts) owns the live workbook state and is the right place to manage recalculation.
- [`src/control-server.ts`](/c:/work/spready/src/control-server.ts) is the TCP surface for external clients.
- [`src/mcp-stdio.ts`](/c:/work/spready/src/mcp-stdio.ts) is the typed MCP wrapper over the TCP surface.
- [`src/App.tsx`](/c:/work/spready/src/App.tsx) is only a view over workbook state and should stay that way.

That means formulas should be implemented as a workbook capability, not as a renderer-only behavior.

## Core Design Decision

Keep raw cell storage unchanged in the workbook model.

`WorkbookSheet.cells: string[][]` should continue to store the raw input string:

- `"123"` remains a literal numeric-looking string.
- `"hello"` remains literal text.
- `"=A1+B1"` is a formula because it begins with `=`.
- `" =A1+B1"` remains literal text in v1 because no implicit trim should change persisted user data.

This keeps the design simple and preserves current transaction and CSV behavior:

- `setCell` and `setRange` still write strings.
- `replaceSheetFromCsv` still imports strings.
- `getSheetCsv` still serializes the raw strings.

The formula system becomes a computed read model layered on top of the stored strings.

## Proposed Modules

### 1. `workbook-core.ts`

Keep it responsible for:

- Workbook structure
- Raw cell storage
- CSV import and export
- Transaction application

Add only small helpers here if needed:

- `isFormulaInput(value: string): boolean`
- `parseCellReference(reference: string): { rowIndex: number; columnIndex: number }`
- shared types for formula-aware reads

Do not move evaluation logic into the renderer.

### 2. New `formula-engine.ts`

Add a new pure module that owns formula parsing and evaluation.

Recommended responsibilities:

- Tokenize formula text after the leading `=`
- Parse the tokens into a small AST
- Evaluate one cell with recursion, memoization, and cycle detection
- Build dependency metadata while evaluating

Suggested internal layers:

1. `tokenizeFormula(input: string): Token[]`
2. `parseFormula(tokens: Token[]): FormulaAst`
3. `evaluateCell(address, context): CellEvaluation`
4. `evaluateSheet(sheet): SheetEvaluationSnapshot`

Suggested grammar:

```text
formula     := "=" expression
expression  := term (("+" | "-") term)*
term        := unary (("*" | "/") unary)*
unary       := ("+" | "-") unary | primary
primary     := number | cellRef | "(" expression ")"
cellRef     := columnLabel rowNumber
columnLabel := /[A-Z]+/
rowNumber   := /[1-9][0-9]*/
```

The parser should ignore whitespace between tokens.

### 3. `workbook-controller.ts`

Make the controller the owner of recalculation and cached computed views.

Recommended private state:

```ts
type CellKey = `${number}:${number}`;

interface CellEvaluation {
  input: string;
  display: string;
  isFormula: boolean;
  numericValue?: number;
  errorCode?: "PARSE" | "REF" | "DIV0" | "VALUE" | "CYCLE";
  dependencies: CellKey[];
}

interface SheetEvaluationSnapshot {
  sheetId: string;
  workbookVersion: number;
  cells: Map<CellKey, CellEvaluation>;
  dependents: Map<CellKey, Set<CellKey>>;
  precedents: Map<CellKey, Set<CellKey>>;
}
```

Pragmatic invalidation rule for v1:

- Any committed write that changes a sheet invalidates that entire sheet's evaluation snapshot.
- The next formula-aware read rebuilds the snapshot for that sheet.

This is intentionally coarse-grained. It is simpler than incremental invalidation and is acceptable for the current workbook size and the limited formula grammar.

## Evaluation Semantics

### Cell value rules

- Non-formula cells display their raw input.
- Formula cells display either the evaluated numeric result or an error marker.
- Empty referenced cells evaluate as numeric `0`.
- Referencing a non-empty text cell in arithmetic returns a value error.

Recommended error display strings:

- parse error: `#ERROR!`
- bad reference: `#REF!`
- divide by zero: `#DIV/0!`
- text in arithmetic: `#VALUE!`
- cycle detected: `#CYCLE!`

### Number formatting

Keep formatting intentionally simple for v1:

- Use JavaScript number evaluation.
- Render integers without a decimal suffix.
- Render decimal results with normal string conversion.
- Do not add locale-specific formatting.

### Reference rules

- References are same-sheet only.
- `A1` means row `0`, column `0`.
- References outside the current sheet bounds return `#REF!`.

## Read Model and API Changes

The main feature requirement creates two parallel read needs:

- Raw input for CSV, formula bar, and precise editing
- Evaluated display values for the spreadsheet grid

Because of that, the read APIs should expose both views instead of overloading one string everywhere.

### New types

Suggested additions:

```ts
export interface CellDataRequest {
  sheetId?: string;
  rowIndex: number;
  columnIndex: number;
}

export interface CellDataResult {
  sheetId: string;
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  input: string;
  display: string;
  isFormula: boolean;
  errorCode?: "PARSE" | "REF" | "DIV0" | "VALUE" | "CYCLE";
}

export interface SheetDisplayRangeResult {
  sheetId: string;
  sheetName: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  values: string[][];
}
```

### Controller methods

Add:

- `getCellData(request: CellDataRequest): CellDataResult`
- `getSheetDisplayRange(request: SheetRangeRequest): SheetDisplayRangeResult`

Keep existing methods unchanged where possible:

- `getSheetRange` continues to return raw input strings for backward compatibility.
- `getSheetCsv` continues to export raw input strings.
- `applyTransaction` stays the only write path.

This preserves existing remote clients while adding explicit formula-aware reads.

## UI Design

The renderer should not evaluate formulas locally. It should request both workbook metadata and formula-aware reads from the main process.

### Formula bar behavior

Add a formula bar above the grid:

- If no cell is selected, show an empty, disabled field.
- If a cell is selected, show the raw `input` value from `getCellData`.
- Editing the formula bar writes back through the existing `applyTransaction` path with `setCell`.

### Grid behavior

The grid should render evaluated display values from `getSheetDisplayRange`.

For a formula cell:

- Grid cell content shows the evaluated result.
- Formula bar shows the raw formula, for example `=A1+B1`.

### Selection flow

Recommended renderer state:

- `selectedCell`
- `selectedCellData`
- `displayRangeCache`

Recommended fetch behavior:

1. Visible region changes trigger `getSheetDisplayRange`.
2. Selected cell changes trigger `getCellData`.
3. Workbook version changes invalidate both caches and refetch as needed.

## TCP Control Server Changes

Add the same formula-aware reads to the TCP protocol in [`src/control-server.ts`](/c:/work/spready/src/control-server.ts):

- `getCellData`
- `getSheetDisplayRange`

Update `listMethods` accordingly.

`applyTransaction` does not need a new formula-specific write method because formulas are just strings written through existing transaction operations.

Recommended examples:

```json
{"id":7,"method":"getCellData","params":{"sheetId":"sheet-1","rowIndex":1,"columnIndex":2}}
```

```json
{"id":8,"method":"getSheetDisplayRange","params":{"sheetId":"sheet-1","startRow":0,"startColumn":0,"rowCount":20,"columnCount":8}}
```

## MCP Changes

Mirror the TCP additions in [`src/mcp-stdio.ts`](/c:/work/spready/src/mcp-stdio.ts):

- `get_cell_data`
- `get_sheet_display_range`

Also update:

- `describe_capabilities`
- `spready://guide`
- the documented transaction and read conventions

The important rule is that formula support must be fully reachable through MCP, not only through the UI.

## Structural Edit Policy

For the first release, keep structural edits simple:

- `insertRows`
- `deleteRows`
- `insertColumns`
- `deleteColumns`
- `resizeSheet`

These operations continue to mutate the raw cell matrix only.

Formulas are recalculated from whatever raw formula strings remain after the structural edit. The formula text itself is not rewritten in v1.

This should be documented clearly because it differs from mature spreadsheet products. It is acceptable for a very basic first implementation and avoids mixing formula parsing logic into transaction rewriting.

## Error Handling

Errors should remain cell-local and non-fatal:

- A bad formula should not prevent unrelated cells from evaluating.
- A cycle should mark only the participating cells with `#CYCLE!`.
- TCP and MCP requests should still succeed unless the request itself is invalid.

This means formula evaluation errors belong in `CellDataResult` and display ranges, not as transport-level failures.

## Implementation Sequence

### Phase 1. Formula engine

- Add tokenization, parsing, and evaluation in a pure module.
- Add unit tests for precedence, parentheses, unary minus, references, divide by zero, invalid references, and cycles.

### Phase 2. Controller integration

- Add lazy per-sheet evaluation snapshots in `WorkbookController`.
- Invalidate the relevant sheet snapshot after any committed workbook change.
- Expose `getCellData` and `getSheetDisplayRange`.

### Phase 3. Transport sync

- Add matching TCP methods.
- Add matching MCP tools and schemas.
- Update MCP guide text and capability description.

### Phase 4. Renderer

- Add selected-cell state.
- Add a formula bar bound to raw input.
- Switch the grid data source from raw range values to evaluated display range values.

### Phase 5. Validation

Add tests at these levels:

- formula engine unit tests
- workbook controller tests for raw versus display reads
- TCP server method coverage
- MCP schema and tool coverage where practical

## Testing Recommendations

At minimum, add tests for these cases:

- `=1+2*3` returns `7`
- `=(1+2)*3` returns `9`
- `=A1+B1` resolves cell references
- empty referenced cells behave as `0`
- `=A1/0` returns `#DIV/0!`
- `=A1+B1` returns `#VALUE!` when `A1` is text
- `=A1` in `A1` returns `#CYCLE!`
- CSV import and export preserve `=A1+B1` as text
- `getSheetRange` returns raw input while `getSheetDisplayRange` returns evaluated display
- `getCellData` returns both the raw formula and the displayed result

## Summary

The simplest durable design is:

- Keep raw strings as the only persisted workbook cell format.
- Add one formula engine module for parse and evaluate.
- Let the controller own recalculation and computed caches.
- Expose explicit raw and display read APIs.
- Mirror those APIs through TCP and MCP.
- Keep the renderer thin, with the formula bar showing raw input and the grid showing evaluated output.

That fits the current Spready architecture, keeps CSV behavior stable, and leaves room for future spreadsheet features without forcing a rewrite of the current workbook model.
