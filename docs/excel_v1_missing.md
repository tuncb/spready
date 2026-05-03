# Excel V1 Missing Features

## Purpose

This document summarizes the gaps between `docs/excel_spec.md` and the current Spready formula implementation, with extra investigation notes for the next likely Excel v1 candidates:

- reference union and intersection
- `VLOOKUP`
- optional date/time functions

The current implementation keeps workbook truth in the main process. Formula support is implemented as a computed read model in `src/formula-engine.ts`, surfaced through the controller, TCP, and MCP display-read APIs.

## Current Formula Gaps

The current evaluator already supports A1 references, same-sheet and cross-sheet ranges with `:`, arithmetic, comparisons, text concatenation, common error literals, and a practical set of functions including `INDEX`, `MATCH`, and `XLOOKUP`.

The main missing or incomplete Excel v1 items are:

- Absolute and mixed references: `$A$1`, `A$1`, and `$A1` are not parsed.
- Defined names: unknown names currently evaluate to `#NAME?`; there is no workbook or sheet defined-name model.
- `LET`: local formula names are not implemented.
- Reference union/intersection: comma-as-union and space-as-intersection are not implemented.
- `VLOOKUP`: not implemented. The minimal checklist allows `XLOOKUP` or `VLOOKUP`, and Spready currently implements `XLOOKUP`.
- Date/time functions: `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, and `DAY` are not implemented.
- Excel worksheet limits: formulas evaluate against current Spready sheet bounds, not Excel's `XFD` and `1048576` maximums.
- Excel numeric precision and formula length limits: the evaluator does not enforce Excel's 15-digit precision or 8192-character formula-content limit.
- Formula reference rewriting: references are not rewritten during copy/paste, row insert/delete, or column insert/delete.

## Reference Union And Intersection

### Current State

Colon ranges already work, for example `A1:B5`. They are represented by a rectangular `RangeValue` with a `cells: CellAddress[][]` shape.

Whitespace is currently discarded during tokenization. That means `SUM(B7:D7 C6:C8)` loses the only syntax signal for reference intersection before parsing. Commas are tokenized, but they are only consumed as function argument separators.

Relevant implementation points:

- `tokenizeFormula` skips whitespace.
- `parseFunctionCall` consumes comma tokens between function arguments.
- `parseRange` only parses the `:` reference operator.
- `createRangeValue` builds a single rectangular range.
- `flattenRangeCells`, aggregate functions, `TEXTJOIN`, and lookup helpers assume one range shape.

### Feasibility

Intersection is medium-high feasibility if kept to rectangular references. The existing `RangeValue` can represent a rectangular intersection result. Empty intersections can naturally map to `#NULL!`, which already exists as an error literal and display value.

Union is medium feasibility and more invasive. A true Excel union can contain multiple disjoint areas. The current `RangeValue` cannot represent this directly without either:

- changing range values to support multiple areas, or
- flattening union results eagerly into a list of cells.

Changing the value model is cleaner, but it touches every helper that consumes ranges.

### Implementation Direction

For intersection:

- Preserve enough whitespace information in the tokenizer to identify single-space intersection between reference expressions.
- Add a parser level for reference operators above unary/scalar operators.
- Evaluate intersection by computing the overlap rectangle between two range/reference values.
- Return `#NULL!` when the overlap is empty.

For union:

- Introduce a multi-area reference value or extend `RangeValue`.
- Keep function argument comma behavior distinct from comma-as-union. A practical first implementation can support union only in parenthesized reference expressions, for example `SUM((A1:A2,C1:C2))`, if broad comma ambiguity becomes too risky.
- Update range consumers to iterate over all areas.

### Corner Cases

- Ordinary whitespace around `+`, `-`, function calls, and parentheses must remain insignificant.
- `SUM(A1:A2,C1:C2)` currently means two arguments. It should not silently become one union argument unless parenthesized or deliberately specified.
- Empty intersections should return `#NULL!`.
- Cross-sheet intersection should probably return `#NULL!`, because ranges on different sheets cannot overlap.
- Cross-sheet union can be allowed if the value model supports multiple areas with sheet ids.
- Overlapping union ranges need a policy: count duplicates by area like Excel-style union behavior, or deduplicate cell addresses.
- Dependencies should reflect the cells actually needed. For intersection, recording the full input ranges before overlap may overstate precedents.

## VLOOKUP

### Current State

`VLOOKUP` is not registered in the formula function registry. Related infrastructure already exists:

- `getRangeArgument` validates range arguments.
- `getVectorAddresses` supports one-dimensional lookup vectors.
- `compareScalarValues` provides current comparison behavior.
- `evaluateXLookup` implements exact lookup across separate lookup and return vectors.
- `evaluateMatch` implements exact and approximate match modes.

### Feasibility

`VLOOKUP` is relatively low-risk and contained. It should be implemented entirely in `src/formula-engine.ts` as a new function handler plus a registry entry. No new workbook-core transaction, TCP method, or MCP tool schema is needed.

MCP guide text should be updated after implementation so external clients know `VLOOKUP` is supported.

### Suggested Behavior

Signature:

```text
VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

Practical first version:

- Require 3 or 4 arguments.
- Require `table_array` to be a rectangular range.
- Treat `col_index_num` as 1-based.
- Use the first column of `table_array` as the lookup vector.
- Return the value from the matching row and requested table column.
- Support exact lookup when `range_lookup` is omitted or false if Spready wants consistency with `XLOOKUP`.
- Alternatively, match Excel more closely by defaulting omitted `range_lookup` to approximate true.

The default for omitted `range_lookup` is the main product decision. Excel defaults to approximate lookup. Spready's existing lookup behavior is more exact-first because `XLOOKUP` defaults to exact.

### Corner Cases

- `col_index_num < 1` should likely return `#VALUE!`.
- `col_index_num > table width` should return `#REF!`.
- A missing exact match should return `#N/A`.
- Errors in the lookup column should propagate.
- Errors in the matched return cell should propagate.
- Text matching should follow current case-insensitive comparison behavior.
- Approximate lookup over unsorted data needs an explicit policy. Excel expects sorted data, while the current `MATCH` implementation scans for the best candidate.
- Cross-sheet table ranges should work through the existing range/address model.

## Date And Time Functions

### Current State

`TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, and `DAY` are not registered in the function registry.

The current formula value model already supports numbers, and `docs/excel_spec.md` recommends treating dates/times as numbers with formatting handled outside the formula language. That fits Spready's current display model, where numbers are rendered by simple string conversion.

### Feasibility

`DATE`, `YEAR`, `MONTH`, and `DAY` are low to medium risk once the serial-date policy is chosen.

`TODAY` and `NOW` are medium risk because they are volatile. The controller caches evaluation snapshots by workbook version, so these functions would otherwise remain fixed until the workbook changes or the cache is explicitly invalidated.

### Implementation Direction

For non-volatile date functions:

- Add serial-date conversion helpers in `src/formula-engine.ts`.
- Add handlers for `DATE`, `YEAR`, `MONTH`, and `DAY`.
- Register the functions in the function registry.
- Add formula-engine tests for serial conversion, overflow handling, and extraction.

For volatile functions:

- Add an evaluation clock dependency so tests can use a stable date/time.
- Decide how volatile recalculation should work with `WorkbookController` caching.
- Consider a volatile flag in evaluation snapshots or a short cache TTL for snapshots containing volatile formulas.

### Serial-Date Policy

The main design choice is whether to emulate Excel's 1900 date system exactly, including the historical leap-year bug, or use a simpler serial model.

Options:

- Excel-compatible 1900 model: better compatibility, but has special-case behavior around serial 60.
- Simple JavaScript-based model: easier to reason about, but may differ from Excel for older dates.

Because this document targets Excel v1 compatibility, the Excel-compatible 1900 model is probably the better default if the implementation remains compact and well tested.

### Corner Cases

- Local time versus UTC for `TODAY()` and `NOW()`.
- Daylight saving time when converting local dates to serial numbers.
- `DATE(2024,13,1)` and `DATE(2024,1,0)` overflow behavior.
- Invalid or negative serial inputs to `YEAR`, `MONTH`, and `DAY`.
- Fractional serials: `NOW()` should include a fractional day; `TODAY()` should not.
- Display formatting is not implemented, so dates will display as serial numbers.
- Cached volatile values should not become stale across UI, TCP, and MCP display reads.

## Transport And Documentation Impact

These features should remain formula-engine behavior. The renderer, TCP server, and MCP server should stay thin.

Expected transport work after implementation:

- No new TCP methods are required. Existing display reads should show the new results.
- No new MCP tools are required. Existing `get_sheet_display_range` and `get_cell_data` should expose the new results.
- MCP capability and guide text should be updated so automation clients know the newly supported functions/operators.
- Tests should cover formula-engine behavior directly and, where useful, controller display reads to prove the computed view reaches TCP/MCP adapters through existing paths.

## Suggested Implementation Order

1. `VLOOKUP`, because it is useful, contained, and reuses existing lookup helpers.
2. `DATE`, `YEAR`, `MONTH`, and `DAY`, after choosing the serial-date policy.
3. `TODAY` and `NOW`, after deciding volatile cache semantics.
4. Reference intersection, because it requires tokenizer/parser changes but can still use a rectangular range result.
5. Reference union, because it probably requires a multi-area range value and broader range-helper updates.
