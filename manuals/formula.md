# Formula Manual

Spready formulas are Excel-like. Enter a formula by storing raw cell text that
starts with `=`. Raw input is preserved; evaluated values are a display/read
model used by the grid, TCP server, and MCP server.

Formula references use one-based A1 notation. TCP and MCP range coordinates use
zero-based row and column indexes.

## Editing And Automation

- UI: edit the selected cell through the formula bar or grid cell entry.
- TCP: write formula text with `applyTransaction` operations such as `setCell`
  and `setRange`.
- MCP: write formula text with `apply_transaction` operations such as `setCell`
  and `setRange`.
- Raw reads: `getSheetRange` / `get_sheet_range`.
- Evaluated reads: `getSheetDisplayRange` / `get_sheet_display_range`.
- Raw plus evaluated value for one cell: `getCellData` / `get_cell_data`.

## Literals

- Numbers: `12`, `-3.5`, `.25`, `1.2e3`.
- Text: `"hello"`. Escape a quote by doubling it: `"He said ""hi"""`.
- Booleans: `TRUE`, `FALSE`, `TRUE()`, `FALSE()`.
- Errors: `#DIV/0!`, `#NAME?`, `#NULL!`, `#VALUE!`, `#REF!`, `#NUM!`,
  `#N/A`.

## Operators

- Arithmetic: `+`, `-`, `*`, `/`, `^`, unary `+`, unary `-`, postfix `%`.
- Text: `&`.
- Comparison: `=`, `<>`, `<`, `<=`, `>`, `>=`.
- Reference range: `A1:B5`.
- Reference intersection: `A1:C3 B2:D4`.
- Parenthesized reference union: `(A1:A2,C1:C2)`.
- Grouping: `( ... )`.

Precedence, low to high: comparisons, `&`, `+`/`-`, `*`/`/`, `^`, `%`, unary
`+`/`-`, references/literals/functions/parentheses. Whitespace is ignored except
when it separates two reference expressions, where it means intersection.

## Cell And Sheet References

- Same-sheet cells: `A1`, `b12`.
- Same-sheet ranges: `A1:B5`; reversed endpoints such as `B5:A1` are normalized.
- Simple sheet names: `Data!A1`, `Data!A1:B10`.
- Quoted sheet names: `'Data Sheet'!A1`, `'Bob''s Sheet'!A1`.
- Qualified range endpoints: `Data!A1:Data!B10`.

Sheet and cell reference matching is case-insensitive. Range endpoints must
resolve to the same sheet. References outside current sheet bounds return
`#REF!`.

Unsupported: absolute references with `$`, whole-column references such as
`A:A`, whole-row references such as `1:1`, R1C1 references, 3D sheet ranges, and
external workbook links.

## Table References

Supported structured references select table data-body cells:

- Current table column: `[Score]`.
- Current row column: `[@Score]`.
- Named table column: `Table1[Score]`.
- Quoted table name: `'Table 1'[Score]`.
- Explicit data qualifier: `Table1[[#Data],[Score]]`.
- Table column range: `Table1[[Q1]:[Q4]]`.
- Current-row column range: `[@[Q1]:[Q4]]`.
- Current-row qualifier form: `[[#This Row],[Score]]`.

Table names and column names are case-insensitive. If a table has no header row,
columns are named `Column1`, `Column2`, and so on. Blank header cells also use
that fallback. Duplicate matching column names return `#REF!`.

If the table name is omitted, the formula cell must be inside a table. Current
row references must be in the table data body. Header, totals, and all-table
sections such as `#Headers`, `#Totals`, and `#All` are not supported.

Editing behavior: when a transaction writes a formula containing `[@` into a
table data-body cell, Spready fills the same formula through that table body
column as a calculated column.

## Functions

Function names are case-insensitive. Unsupported names return `#NAME?`.

| Category         | Supported functions                                                                                                                                 |
| ---------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| Aggregate/math   | `SUM(values...)`, `PRODUCT(values...)`, `MIN(values...)`, `MAX(values...)`, `AVERAGE(values...)`, `COUNT(values...)`, `COUNTA(values...)`           |
| Numeric          | `ABS(number)`, `ROUND(number,digits)`, `INT(number)`, `MOD(number,divisor)`, `POWER(number,power)`, `SQRT(number)`, `LN(number)`                    |
| Logical          | `AND(values...)`, `OR(values...)`, `NOT(value)`, `IF(test,then,[else])`, `IFERROR(value,fallback)`, `TRUE()`, `FALSE()`                             |
| Text             | `LEN(text)`, `LEFT(text,[count])`, `RIGHT(text,[count])`, `MID(text,start,count)`, `TRIM(text)`, `LOWER(text)`, `UPPER(text)`, `CONCAT(values...)`  |
| Text conversion  | `TEXTJOIN(delimiter,ignore_empty,values...)`, `VALUE(text)`                                                                                         |
| Date/time        | `TODAY()`, `NOW()`, `DATE(year,month,day)`, `YEAR(serial)`, `MONTH(serial)`, `DAY(serial)`                                                          |
| Lookup/reference | `CHOOSE(index,values...)`, `ROW([reference])`, `COLUMN([reference])`, `INDEX(range,row,[column])`                                                   |
| Lookup           | `MATCH(lookup,range,[match_type])`, `XLOOKUP(lookup,lookup_range,return_range,[if_not_found])`, `VLOOKUP(lookup,table_range,column,[range_lookup])` |

## Evaluation Rules

- A cell is a formula only if its raw value starts exactly with `=`.
- A leading space before `=` makes the cell plain text.
- Non-formula raw cells are blank, numeric if the raw text is numeric, or text
  otherwise. Raw `TRUE` and `FALSE` are text unless produced by a formula.
- A single-cell range can be used as a scalar. A multi-cell range used where a
  scalar is required returns `#VALUE!`; there is no spill behavior.
- Aggregate functions include numeric cells in ranges and ignore blank/text/
  boolean range members. Direct scalar arguments are coerced, so invalid direct
  text can return `#VALUE!`.
- `IF` and `IFERROR` evaluate only the branch they need.
- Text comparisons are case-insensitive.
- Dates use Excel 1900-system serial numbers, including serial `60` for
  compatibility. `TODAY()` and `NOW()` use local time and are volatile. Date
  results display as serial numbers unless number formatting changes the display.
- Circular dependencies display as `#CYCLE!`.

## Excel Differences And Limits

- The function set is a practical subset of Excel, not the full worksheet
  catalog.
- Absolute references, defined names, `LET`, dynamic arrays, array constants,
  external workbook links, and full Excel structured-reference syntax are not
  supported.
- `MATCH` defaults to exact match (`0`) instead of Excel's legacy approximate
  default.
- `VLOOKUP` defaults omitted `range_lookup` to exact match. Pass `TRUE` for
  approximate lookup.
- `XLOOKUP` supports exact match plus optional `if_not_found`; other Excel
  `XLOOKUP` modes are not implemented.
- Reference union is supported only as a parenthesized reference expression.
  Unparenthesized commas inside function calls are argument separators.
- Formula strings are not rewritten when rows, columns, or tables move. During
  table sorts, formulas move with their rows; structured references like
  `[@Score]` keep current-row meaning.

## Error Displays

| Display   | Meaning                                                                               |
| --------- | ------------------------------------------------------------------------------------- |
| `#ERROR!` | Parse failure or malformed syntax                                                     |
| `#REF!`   | Invalid reference, table column, range target, or lookup/index target                 |
| `#DIV/0!` | Division by zero or averaging no numeric values                                       |
| `#VALUE!` | Type mismatch, invalid argument count, invalid scalar/range shape, or failed coercion |
| `#CYCLE!` | Circular dependency                                                                   |
| `#NAME?`  | Unknown function or unsupported named reference                                       |
| `#NUM!`   | Invalid numeric/date result                                                           |
| `#N/A`    | Lookup miss or explicit error literal                                                 |
| `#NULL!`  | Empty reference intersection or explicit error literal                                |
