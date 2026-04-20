import type { EChartsOption } from "echarts";
import ReactECharts from "echarts-for-react";

import {
  getColumnTitle,
  type WorkbookChartPreview,
  type WorkbookChartResult,
  type WorkbookChartStatus,
  type WorkbookChartType,
} from "./workbook-core";

export interface WorkbookChartDockEntry {
  chartType: WorkbookChartType;
  id: string;
  name: string;
  status: WorkbookChartStatus;
}

interface WorkbookChartDockProps {
  activeSheetName: string;
  chartEntries: WorkbookChartDockEntry[];
  chartPreview: WorkbookChartPreview | null;
  chartResult: WorkbookChartResult | null;
  isDetailLoading: boolean;
  isListLoading: boolean;
  onSelectChart: (chartId: string) => void;
  selectedChartId: string | null;
}

export function WorkbookChartDock({
  activeSheetName,
  chartEntries,
  chartPreview,
  chartResult,
  isDetailLoading,
  isListLoading,
  onSelectChart,
  selectedChartId,
}: WorkbookChartDockProps) {
  const selectedEntry =
    chartEntries.find((chart) => chart.id === selectedChartId) ?? null;
  const sourceRange =
    chartResult &&
    formatChartRange(chartResult.chart.spec.source.range.startColumn, {
      columnCount: chartResult.chart.spec.source.range.columnCount,
      rowCount: chartResult.chart.spec.source.range.rowCount,
      startRow: chartResult.chart.spec.source.range.startRow,
    });
  const option = chartPreview?.option as EChartsOption | undefined;

  return (
    <aside className="chart-dock" aria-label="Chart preview dock">
      <header className="chart-dock__header">
        <div>
          <p className="chart-dock__eyebrow">Charts</p>
          <h2 className="chart-dock__title">Sheet Graphs</h2>
          <p className="chart-dock__subtitle">
            {activeSheetName}
            {" · "}
            {chartEntries.length === 0
              ? "No charts on this sheet"
              : `${chartEntries.length} chart${
                  chartEntries.length === 1 ? "" : "s"
                } available`}
          </p>
        </div>
        <div className="chart-dock__count">{chartEntries.length}</div>
      </header>

      <div className="chart-dock__body">
        <section
          className="chart-dock__list-section"
          aria-label="Sheet chart list"
        >
          <div className="chart-dock__section-heading">
            <span>Available charts</span>
            {isListLoading ? (
              <span className="chart-dock__muted">Syncing</span>
            ) : null}
          </div>

          <div className="chart-dock__list">
            {chartEntries.length === 0 ? (
              <div className="chart-dock__empty">
                Charts saved in `.spready` files appear here for preview and
                inspection.
              </div>
            ) : (
              chartEntries.map((chart) => (
                <button
                  key={chart.id}
                  className={`chart-dock__list-item${
                    chart.id === selectedChartId ? " is-selected" : ""
                  }`}
                  onClick={() => {
                    onSelectChart(chart.id);
                  }}
                  type="button"
                >
                  <div className="chart-dock__list-copy">
                    <strong>{chart.name}</strong>
                    <span>{formatChartTypeLabel(chart.chartType)}</span>
                  </div>
                  <span
                    className={`chart-dock__status-pill is-${chart.status}`}
                  >
                    {chart.status}
                  </span>
                </button>
              ))
            )}
          </div>
        </section>

        <section className="chart-dock__detail" aria-live="polite">
          {chartEntries.length === 0 ? (
            <div className="chart-dock__empty chart-dock__empty--detail">
              Select a sheet that already contains charts to preview it in the
              desktop app.
            </div>
          ) : isDetailLoading ||
            !selectedEntry ||
            !chartResult ||
            !chartPreview ? (
            <div className="chart-dock__loading">
              Loading chart preview from the workbook controller.
            </div>
          ) : (
            <>
              <div className="chart-dock__detail-header">
                <div>
                  <p className="chart-dock__eyebrow">Preview</p>
                  <h3 className="chart-dock__detail-title">
                    {chartResult.chart.name}
                  </h3>
                </div>
                <span
                  className={`chart-dock__status-pill is-${chartPreview.status}`}
                >
                  {chartPreview.status}
                </span>
              </div>

              <div className="chart-dock__meta-grid">
                <div className="chart-dock__meta-card">
                  <span className="chart-dock__meta-label">Type</span>
                  <strong>
                    {formatChartTypeLabel(selectedEntry.chartType)}
                  </strong>
                </div>
                <div className="chart-dock__meta-card">
                  <span className="chart-dock__meta-label">Source</span>
                  <strong>{sourceRange}</strong>
                </div>
                <div className="chart-dock__meta-card">
                  <span className="chart-dock__meta-label">Layout</span>
                  <strong>
                    {chartResult.chart.spec.source.seriesLayoutBy}
                  </strong>
                </div>
                <div className="chart-dock__meta-card">
                  <span className="chart-dock__meta-label">Dataset</span>
                  <strong>
                    {chartPreview.dataset.dimensions.length} dimensions
                  </strong>
                </div>
              </div>

              {chartPreview.status === "ok" && option ? (
                <div className="chart-dock__preview-frame">
                  <ReactECharts
                    className="chart-dock__echarts"
                    lazyUpdate
                    notMerge
                    option={option}
                  />
                </div>
              ) : (
                <div className="chart-dock__callout chart-dock__callout--invalid">
                  This chart cannot be previewed until its source range and
                  dimensions are valid again.
                </div>
              )}

              {chartPreview.validationIssues.length > 0 ? (
                <div className="chart-dock__callout chart-dock__callout--invalid">
                  <strong>Validation issues</strong>
                  <ul className="chart-dock__issue-list">
                    {chartPreview.validationIssues.map((issue) => (
                      <li key={`${issue.code}:${issue.message}`}>
                        <span>{issue.code}</span>
                        <span>{issue.message}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              ) : null}

              {chartPreview.warnings.length > 0 ? (
                <div className="chart-dock__callout">
                  <strong>Preview warnings</strong>
                  <ul className="chart-dock__issue-list">
                    {chartPreview.warnings.map((warning) => (
                      <li key={warning}>{warning}</li>
                    ))}
                  </ul>
                </div>
              ) : null}
            </>
          )}
        </section>
      </div>
    </aside>
  );
}

function formatChartRange(
  startColumn: number,
  range: {
    columnCount: number;
    rowCount: number;
    startRow: number;
  },
): string {
  if (range.columnCount < 1 || range.rowCount < 1) {
    return "Empty range";
  }

  const endColumn = startColumn + range.columnCount - 1;
  const endRow = range.startRow + range.rowCount;

  return `${getColumnTitle(startColumn)}${range.startRow + 1}:${getColumnTitle(
    endColumn,
  )}${endRow}`;
}

function formatChartTypeLabel(chartType: WorkbookChartType): string {
  switch (chartType) {
    case "area":
      return "Area";
    case "bar":
      return "Bar";
    case "line":
      return "Line";
    case "pie":
      return "Pie";
    case "scatter":
      return "Scatter";
  }
}
