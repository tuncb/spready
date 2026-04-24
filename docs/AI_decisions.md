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
- Native workbook files use `.spready` document format version 3 for chart specs and embedded chart layout persistence. Older workbook document versions are intentionally rejected rather than migrated.
- Controller-side chart reads return shared chart definitions plus derived status and validation issues; chart preview is a read-only projection that adds a normalized dataset and derived ECharts option.
- Preview generation normalizes `seriesLayoutBy: "row"` sources into a column-oriented dataset before building ECharts options, so renderer consumers only need one dataset shape in v1.
- Preview generation converts blank cells and formula errors to `null` points. Formula errors also emit preview warnings instead of failing the entire chart preview.
- Invalid charts still return a preview payload with an empty dataset and a minimal title-only option, rather than throwing from the read path.
- Exact persisted chart writes use transaction operations inside `applyWorkbookTransaction` (`addChart`, `renameChart`, `setChartSpec`, `setChartLayout`, `deleteChart`). The simplified `createChart` convenience method expands to an `addChart` transaction operation rather than implementing separate chart mutation rules.
- TCP chart support mirrors the controller method names and shared payloads directly: `getSheetCharts`, `getSheetChartPreviews`, `getChart`, `getChartPreview`, and the convenience write method `createChart`. Transport layers do not own chart business rules.
- MCP chart support stays a thin adapter over TCP with `get_sheet_charts`, `get_sheet_chart_previews`, `get_chart`, `get_chart_preview`, and `create_chart`. MCP does not implement workbook or chart rules itself.
- Common chart creation uses the shared `CreateChartRequest` contract in `workbook-core.ts` so LLM clients can omit source range, dimensions, headers, and layout when defaults are sufficient. Exact persisted chart edits still use transaction operations.
- Structural row and column edits rewrite persisted chart source ranges inside `applyWorkbookTransaction`, not in the renderer or transport layers. Deleting a sheet keeps its charts as explicit invalid records rather than silently removing them.
- Embedded chart layout is workbook-owned and persisted on each chart as a cell anchor plus pixel offsets, pixel width/height, and z-index. The renderer may keep transient drag/resize state, but committed chart moves and resizes use `setChartLayout` transactions.
- Chart layout anchors move when rows or columns are inserted or deleted before the anchor. Chart sizes remain pixel-based and do not automatically resize with cell structural edits. When sheet dimensions shrink, chart anchors clamp to the remaining sheet bounds so embedded charts stay reachable.
- Embedded chart UI renders charts over the spreadsheet grid rather than in a separate chart pane. ECharts options remain derived preview data; the persisted chart contract stores only Spready chart specs and layout.
- When an embedded chart is selected, the Delete key and delete menu action delete that chart immediately via the existing `deleteChart` transaction. Formula-bar text deletion still takes precedence while the formula input is focused.
- Valid embedded chart previews omit the ECharts `title` option because the overlay header owns the visible chart name. This avoids duplicate titles while keeping the chart name in the persisted chart contract and summary data.
- Valid cartesian chart previews derive x-axis labels from the category dimension and y-axis labels from the single value dimension. Multi-series cartesian charts use `Value` as the y-axis label. The generated cartesian option reserves left grid space and rotates the y-axis name so it remains visible in compact embedded charts. Pie previews use the value dimension as the series label and show slice labels as name plus percent.
- Embedded chart move and resize commits optimistically update the derived sheet preview state before the controller transaction resolves. The controller remains authoritative; failed commits roll back the optimistic preview unless a newer preview has already replaced it.
- The GUI insert-chart action passes the current grid cell selection into the chart editor create request. The editor uses that selected range as the default source range and falls back to the sheet used range when no cell range is selected.
