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

## 2026-04-20

- Chart support is being implemented in the planned rollout order, and each rollout step must finish with formatter, tests, lint, and typecheck before the next step starts.
- The chart stack is Apache ECharts 6 with `echarts-for-react` 3.x. Backward-compatibility with older ECharts releases is intentionally out of scope.
- Workbook chart definitions are modeled as a typed Spready contract in `workbook-core.ts`; renderer-specific ECharts options are a derived view, not the persisted source of truth.
- V1 chart bindings are same-sheet only. A chart's owning `sheetId` and its source range `sheetId` must match.
- V1 chart definitions use a tabular source model aligned with ECharts `dataset` and `encode`, with `seriesLayoutBy` captured explicitly in the shared contract.
- V1 chart dimensions are interpreted against the layout axis chosen by `seriesLayoutBy`: `column` uses source columns as dimensions, and `row` uses source rows as dimensions.
- V1 chart creation must require a non-empty source range and valid dimension indexes. Structural sheet edits may later move or shrink those ranges, and invalid charts should remain explicit rather than being silently deleted.
- Native workbook files move to `.spready` document format version 2 for chart persistence. Older workbook document versions are intentionally rejected rather than migrated.
- Controller-side chart reads return shared chart definitions plus derived status and validation issues; chart preview is a read-only projection that adds a normalized dataset and derived ECharts option.
- Preview generation normalizes `seriesLayoutBy: "row"` sources into a column-oriented dataset before building ECharts options, so renderer consumers only need one dataset shape in v1.
- Preview generation converts blank cells and formula errors to `null` points. Formula errors also emit preview warnings instead of failing the entire chart preview.
- Invalid charts still return a preview payload with an empty dataset and a minimal title-only option, rather than throwing from the read path.
- TCP chart support mirrors the controller method names and shared payloads directly: `getSheetCharts`, `getChart`, and `getChartPreview`. No chart-specific transport contract is introduced.
- MCP chart support stays a thin read-only adapter over TCP with `get_sheet_charts`, `get_chart`, and `get_chart_preview`. MCP does not implement workbook or chart rules itself.
- Structural row and column edits rewrite persisted chart source ranges inside `applyWorkbookTransaction`, not in the renderer or transport layers. Deleting a sheet keeps its charts as explicit invalid records rather than silently removing them.
