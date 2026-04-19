# AI decisions

## 2026-04-19

- Formula-compatibility expansion is being implemented in five sequential tasks, with formatter, tests, lint, and typecheck run between tasks before starting the next one.
- Absolute references such as `$A$1`, `A$1`, and `$A1` are intentionally out of scope for this round. Formulas using `$` should remain unsupported.
- Cross-sheet references, defined names, `LET`, reference union/intersection operators, and formula rewriting during structural edits remain out of scope for this round.
- Raw non-formula cells continue to store plain strings. During evaluation, only numeric-looking raw strings are inferred as numbers; raw `TRUE` and `FALSE` cell contents remain text unless they come from formula evaluation.
- The public workbook/controller/TCP/MCP read APIs stay unchanged. Formula compatibility improvements should flow through the existing raw-vs-display read paths instead of introducing new transport methods.
- Multi-cell ranges are supported as function/operator inputs, but there is no spill or implicit-intersection behavior in this round. A bare multi-cell range used where a scalar is required should evaluate to `#VALUE!`.
- Text comparisons are implemented case-insensitively to align more closely with Excel-style formula behavior, while raw cell display remains unchanged.
- Numeric aggregation functions flatten same-sheet ranges, ignore blank cells, and ignore non-numeric range members. Direct scalar arguments still use normal scalar coercion.
- `IF` and `IFERROR` are evaluated lazily so unused branches do not trigger errors.
- The lookup slice implements `XLOOKUP` instead of `VLOOKUP`, and keeps lookup behavior same-sheet only.
- `MATCH` defaults to exact-match mode in this app instead of Excel’s legacy approximate default. Explicit `1` and `-1` modes are still supported for basic approximate matching.
