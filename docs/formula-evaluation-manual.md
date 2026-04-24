# Formula Evaluation Manual

## Purpose

This manual describes how formula evaluation currently works in Spready across the app UI, the TCP control server, and the MCP wrapper.

The key design rule is:

- The workbook stores the raw cell input exactly as written.
- Formula evaluation is a computed read model over that raw input.
- The main process owns evaluation behavior, so UI, TCP, and MCP all see the same results.

## Where Evaluation Appears

Formula evaluation is exposed consistently across the product:

- The grid uses evaluated display values.
- `get_sheet_display_range` returns evaluated display values.
- `get_cell_data` returns both the raw input and the evaluated display value.
- `get_sheet_range` returns the raw stored input and does not rewrite formulas.

Example:

- Raw cell input: `=SUM(A1:A2)`
- `get_sheet_range`: returns `=SUM(A1:A2)`
- `get_sheet_display_range`: returns the computed result, such as `30`

## Recognition And Storage

A cell is treated as a formula only when its stored string begins with `=`.

Important consequences:

- `=A1+B1` is a formula.
- ` =A1+B1` is plain text because the leading space is preserved.
- Raw input is never normalized into a different persisted value during evaluation.

Non-formula cells are interpreted as:

- blank when the stored string is empty
- number when the raw text is a numeric literal
- text otherwise

Numeric literals support decimals and scientific notation, for example:

- `12`
- `-3.5`
- `.25`
- `1.2e3`

## Supported Formula Syntax

### References

- Same-sheet A1 references are supported.
- References are case-insensitive, so `A1` and `a1` behave the same.
- Multi-cell ranges such as `A1:B5` are supported.
- Range endpoints are normalized, so `B5:A1` evaluates as the same rectangle.

### Literals

Supported literal kinds inside formulas:

- numbers
- text in double quotes
- booleans: `TRUE`, `FALSE`
- error literals: `#DIV/0!`, `#NAME?`, `#NULL!`, `#VALUE!`, `#REF!`, `#NUM!`, `#N/A`

String literals use doubled quotes to escape quotes inside the string.

Example:

```text
="He said ""hello"""
```

### Operators

Supported operators:

- arithmetic: `+`, `-`, `*`, `/`, `^`, `%`
- text concatenation: `&`
- comparison: `=`, `<>`, `<`, `<=`, `>`, `>=`
- range construction: `:`
- parentheses: `(` and `)`

### Operator Precedence

The evaluator currently parses operators in this order, from lowest precedence to highest:

1. comparisons: `=`, `<>`, `<`, `<=`, `>`, `>=`
2. concatenation: `&`
3. addition and subtraction: `+`, `-`
4. multiplication and division: `*`, `/`
5. exponentiation: `^`
6. postfix percent: `%`
7. unary plus and minus: `+`, `-`
8. references, literals, function calls, and parenthesized expressions

Notes:

- Parentheses override normal precedence.
- Repeated exponentiation currently evaluates left-to-right.
- Multi-cell ranges are first-class values inside formulas, but outside functions they only behave like scalars when the range is exactly one cell.
- A multi-cell range used where a scalar is required returns `#VALUE!`.

## Value Types And Display Rules

Internally, evaluated formula results can be:

- blank
- boolean
- number
- text
- error
- range

Display formatting is intentionally simple:

- blank displays as an empty string
- booleans display as `TRUE` or `FALSE`
- numbers display through JavaScript string conversion
- errors display as spreadsheet-style error markers such as `#REF!`

`-0` is normalized to `0` for display.

## Coercion Rules

### Numeric Coercion

When a formula operation expects a number:

- blank becomes `0`
- `TRUE` becomes `1`
- `FALSE` becomes `0`
- numeric text such as `"12.5"` is parsed as a number
- nonnumeric text returns `#VALUE!`

### Text Coercion

When a formula operation expects text:

- blank becomes `""`
- booleans become `TRUE` or `FALSE`
- numbers become their display string

### Boolean Coercion

When a formula operation expects a boolean:

- blank becomes `FALSE`
- zero becomes `FALSE`
- any nonzero number becomes `TRUE`
- text is trimmed and accepts `TRUE`, `FALSE`, or numeric text
- any other text returns `#VALUE!`

### Comparison Rules

Comparison behavior is intentionally simple and predictable:

- blank is coerced to the type of the other operand when possible
- text comparison is case-insensitive
- if both sides can be interpreted as numbers, they are compared numerically
- otherwise values are compared as normalized text

Example:

```text
="A"="a"     -> TRUE
=1<>2        -> TRUE
```

## Supported Functions

### Aggregate And Numeric

- `SUM`
- `PRODUCT`
- `MIN`
- `MAX`
- `AVERAGE`
- `COUNT`
- `COUNTA`
- `ABS`
- `ROUND`
- `INT`
- `MOD`
- `POWER`
- `SQRT`

Notes:

- `AVERAGE` returns `#DIV/0!` when there are no numeric values.
- `SQRT` of a negative number returns `#NUM!`.
- `POWER` returns `#NUM!` when the JavaScript result is `NaN`.

### Logical

- `TRUE()`
- `FALSE()`
- `AND`
- `OR`
- `NOT`
- `IF`
- `IFERROR`

Notes:

- `IF` and `IFERROR` are evaluated lazily, so only the needed branch is evaluated.
- `IF(test, whenTrue)` returns `FALSE` when `test` is false and no third argument is provided.

### Text

- `LEN`
- `LEFT`
- `RIGHT`
- `MID`
- `TRIM`
- `LOWER`
- `UPPER`
- `CONCAT`
- `TEXTJOIN`
- `VALUE`

Notes:

- `TRIM` removes leading and trailing spaces and collapses repeated internal spaces to one space.
- `TEXTJOIN(delimiter, ignoreEmpty, ...)` accepts ranges and scalars.
- `VALUE` parses numeric text and returns `#VALUE!` for nonnumeric text.

### Lookup And Reference

- `CHOOSE`
- `ROW`
- `COLUMN`
- `INDEX`
- `MATCH`
- `XLOOKUP`

Notes:

- `ROW()` and `COLUMN()` without arguments use the current formula cell.
- `ROW(range)` and `COLUMN(range)` return the row or column of the first cell in the range.
- `INDEX(range, row, column?)` supports one-dimensional and two-dimensional same-sheet ranges.
- `MATCH(value, lookupRange, matchType?)` supports `0`, `1`, and `-1`.
- `MATCH` requires a one-dimensional lookup range.
- `XLOOKUP(lookup, lookupRange, returnRange, notFound?)` performs same-sheet lookup over one-dimensional ranges of equal length.
- `XLOOKUP` currently behaves as exact match lookup.

## Aggregate Function Range Behavior

Aggregate functions do not treat direct scalar arguments and range arguments exactly the same.

For `SUM`, `PRODUCT`, `MIN`, `MAX`, `AVERAGE`, and `COUNT`:

- numeric cells inside ranges are included
- blank cells inside ranges are ignored
- text and booleans inside ranges are ignored unless the function specifically works on all nonblank values like `COUNTA`
- direct scalar arguments are coerced, so a direct boolean can count as `1` or `0`
- direct nonnumeric text arguments can produce `#VALUE!`

This is why range-based aggregation is generally safer when mixed user-entered data is expected.

## Error Results

The evaluator can return these display errors:

| Display   | Meaning                                                                               |
| --------- | ------------------------------------------------------------------------------------- |
| `#ERROR!` | Parse failure or malformed formula syntax                                             |
| `#REF!`   | Invalid reference, invalid range target, or out-of-bounds lookup/index target         |
| `#DIV/0!` | Division by zero or averaging no numeric values                                       |
| `#VALUE!` | Type mismatch, invalid argument count, invalid scalar/range shape, or failed coercion |
| `#CYCLE!` | Circular dependency between cells                                                     |
| `#NAME?`  | Unknown function name or unsupported named reference                                  |
| `#NUM!`   | Invalid numeric result, such as `SQRT(-1)`                                            |
| `#N/A`    | Lookup miss or explicit `#N/A` literal                                                |
| `#NULL!`  | Supported as an error literal                                                         |

## Dependency And Cycle Behavior

Evaluation tracks direct precedents and dependents for formula cells.

That means the engine can:

- understand which cells a formula directly depends on
- report which cells depend on a given cell
- detect circular references and surface `#CYCLE!`

Cycles are reported at display time and remain part of the evaluated view until the underlying raw input is changed.

## Current Exclusions

The current evaluator intentionally does not support:

- absolute references with `$`
- cross-sheet references
- defined names
- `LET`

Also keep these implementation boundaries in mind:

- formula evaluation is same-sheet only
- there is no separate array formula model
- multi-cell ranges are not implicitly spilled into neighboring cells
- unknown identifiers resolve to `#NAME?`

## Practical Examples

### Basic Arithmetic

```text
=1+2*3          -> 7
=(1+2)*3        -> 9
=-(A1+B1)       -> unary minus over a grouped expression
```

### Text And Logic

```text
="Hello" & " " & "World"        -> Hello World
=IF(B1=0,"zero",A1/B1)          -> lazy branch evaluation
=IFERROR(A1/B1,99)              -> fallback on error
```

### Lookup

```text
=MATCH("b",A1:A3)                       -> position in a one-dimensional range
=INDEX(B1:B3,2)                         -> second value from a column range
=XLOOKUP("c",A1:A3,B1:B3,"nf")          -> exact match with fallback
```

### Raw Versus Display

If a cell stores `=TEXTJOIN(", ",TRUE,A1:A2)`:

- raw read: `=TEXTJOIN(", ",TRUE,A1:A2)`
- display read: `a, b`

## Guidance For Clients

When building against TCP or MCP:

- use `get_sheet_range` when the raw formula text matters
- use `get_sheet_display_range` when evaluated grid values matter
- use `get_cell_data` when both the raw formula and the computed display must be shown together

This keeps automation aligned with the UI and preserves the architectural rule that workbook truth stays in the main process while transport layers remain thin.
