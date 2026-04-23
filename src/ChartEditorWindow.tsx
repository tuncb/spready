import { useEffect, useMemo, useState } from "react";

import {
  buildChartEditorOperations,
  createChartEditorFormState,
  getChartEditorSheetId,
  getChartEditorValidationIssues,
  type ChartEditorFormState,
  type ChartEditorWindowRequest,
} from "./chart-editor-state";
import { getColumnTitle, type WorkbookChart, type WorkbookSummary } from "./workbook-core";

function getChartEditorRequest(): ChartEditorWindowRequest {
  const params = new URLSearchParams(window.location.search);
  const mode = params.get("mode");

  if (mode === "edit") {
    const chartId = params.get("chartId");

    if (!chartId) {
      throw new Error("Chart editor requires a chartId in edit mode.");
    }

    return {
      chartId,
      mode: "edit",
    };
  }

  return {
    mode: "create",
    sheetId: params.get("sheetId") ?? undefined,
  };
}

export function ChartEditorWindow() {
  const request = useMemo(() => getChartEditorRequest(), []);
  const [chart, setChart] = useState<WorkbookChart | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [formState, setFormState] = useState<ChartEditorFormState | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [summary, setSummary] = useState<WorkbookSummary | null>(null);

  useEffect(() => {
    document.title = request.mode === "edit" ? "Edit Chart" : "Create Chart";
  }, [request]);

  useEffect(() => {
    let isCancelled = false;

    const load = async () => {
      setIsLoading(true);
      setErrorMessage(null);

      try {
        const nextSummary = await window.appShell.getWorkbookSummary();
        const nextChart =
          request.mode === "edit"
            ? (await window.appShell.getChart(request.chartId)).chart
            : undefined;
        const sheetId = getChartEditorSheetId(request, nextSummary, nextChart);
        const usedRange = await window.appShell.getUsedRange(sheetId);
        const nextFormState = createChartEditorFormState(
          request,
          nextSummary,
          usedRange,
          nextChart,
        );

        if (isCancelled) {
          return;
        }

        setSummary(nextSummary);
        setChart(nextChart ?? null);
        setFormState(nextFormState);
      } catch (error) {
        if (isCancelled) {
          return;
        }

        setErrorMessage(
          error instanceof Error
            ? error.message
            : "Chart editor could not be loaded.",
        );
      } finally {
        if (!isCancelled) {
          setIsLoading(false);
        }
      }
    };

    void load();

    return () => {
      isCancelled = true;
    };
  }, [request]);

  const sheetId =
    summary && formState
      ? getChartEditorSheetId(request, summary, chart ?? undefined)
      : null;
  const sheetName =
    summary && sheetId
      ? summary.sheets.find((sheet) => sheet.id === sheetId)?.name ?? sheetId
      : "";
  const validationIssues =
    summary && formState && sheetId
      ? getChartEditorValidationIssues(
          formState,
          sheetId,
          summary,
          request.mode === "edit" ? request.chartId : undefined,
        )
      : [];
  const canSave =
    !isLoading &&
    !isSaving &&
    summary !== null &&
    formState !== null &&
    sheetId !== null &&
    validationIssues.length === 0;
  const sourceRange =
    formState === null ? "" : formatRangePreview(formState);

  const updateField = <Key extends keyof ChartEditorFormState>(
    field: Key,
    value: ChartEditorFormState[Key],
  ) => {
    setFormState((current) =>
      current === null
        ? current
        : {
            ...current,
            [field]: value,
          },
    );
  };

  const handleChartTypeChange = (nextChartType: ChartEditorFormState["chartType"]) => {
    setFormState((current) => {
      if (current === null) {
        return current;
      }

      return {
        ...current,
        categoryDimension: nextChartType === "pie" ? current.categoryDimension : current.categoryDimension,
        chartType: nextChartType,
        nameDimension: nextChartType === "pie" ? current.nameDimension : "0",
        smooth:
          nextChartType === "line" || nextChartType === "area"
            ? current.smooth
            : false,
        stacked:
          nextChartType === "bar" || nextChartType === "area"
            ? current.stacked
            : false,
        valueDimension: nextChartType === "pie" ? current.valueDimension : "1",
        valueDimensions:
          nextChartType === "pie" ? current.valueDimensions : current.valueDimensions,
      };
    });
  };

  const handleSubmit = async () => {
    if (!summary || !formState || !sheetId) {
      return;
    }

    setIsSaving(true);
    setErrorMessage(null);

    try {
      await window.appShell.applyTransaction({
        operations: buildChartEditorOperations(request, sheetId, formState),
      });
      window.close();
    } catch (error) {
      setErrorMessage(
        error instanceof Error ? error.message : "Chart could not be saved.",
      );
      setIsSaving(false);
    }
  };

  if (isLoading || formState === null || summary === null || sheetId === null) {
    return (
      <main className="chart-editor-window">
        <section className="chart-editor-panel">
          <div className="chart-editor__loading">
            {errorMessage ?? "Loading chart editor..."}
          </div>
          <footer className="chart-editor__footer">
            <button
              className="chart-editor__button chart-editor__button--secondary"
              onClick={() => {
                window.close();
              }}
              type="button"
            >
              Close
            </button>
          </footer>
        </section>
      </main>
    );
  }

  return (
    <main className="chart-editor-window">
      <section className="chart-editor-panel">
        <header className="chart-editor__header">
          <div>
            <p className="chart-editor__eyebrow">
              {request.mode === "edit" ? "Edit chart" : "Insert chart"}
            </p>
            <h1 className="chart-editor__title">
              {request.mode === "edit" ? "Edit Chart" : "Create Chart"}
            </h1>
            <p className="chart-editor__subtitle">
              {sheetName}
              {" · "}
              {sourceRange}
            </p>
          </div>
        </header>

        <div className="chart-editor__body">
          <section className="chart-editor__section">
            <div className="chart-editor__field">
              <label htmlFor="chart-name">Chart name</label>
              <input
                id="chart-name"
                onChange={(event) => {
                  updateField("name", event.target.value);
                }}
                value={formState.name}
              />
            </div>

            <div className="chart-editor__field">
              <label htmlFor="chart-type">Chart type</label>
              <select
                id="chart-type"
                onChange={(event) => {
                  handleChartTypeChange(
                    event.target.value as ChartEditorFormState["chartType"],
                  );
                }}
                value={formState.chartType}
              >
                <option value="bar">Bar</option>
                <option value="line">Line</option>
                <option value="area">Area</option>
                <option value="scatter">Scatter</option>
                <option value="pie">Pie</option>
              </select>
            </div>
          </section>

          <section className="chart-editor__section">
            <div className="chart-editor__field">
              <label htmlFor="chart-layout">Series layout</label>
              <select
                id="chart-layout"
                onChange={(event) => {
                  updateField(
                    "seriesLayoutBy",
                    event.target.value as ChartEditorFormState["seriesLayoutBy"],
                  );
                }}
                value={formState.seriesLayoutBy}
              >
                <option value="column">By column</option>
                <option value="row">By row</option>
              </select>
            </div>

            <label className="chart-editor__checkbox">
              <input
                checked={formState.sourceHeader}
                onChange={(event) => {
                  updateField("sourceHeader", event.target.checked);
                }}
                type="checkbox"
              />
              <span>First row or column contains headers</span>
            </label>
          </section>

          <section className="chart-editor__section chart-editor__section--grid">
            <h2 className="chart-editor__section-title">Source range</h2>

            <div className="chart-editor__field">
              <label htmlFor="chart-start-row">Start row</label>
              <input
                id="chart-start-row"
                inputMode="numeric"
                onChange={(event) => {
                  updateField("startRow", event.target.value);
                }}
                value={formState.startRow}
              />
            </div>

            <div className="chart-editor__field">
              <label htmlFor="chart-start-column">Start column</label>
              <input
                id="chart-start-column"
                inputMode="numeric"
                onChange={(event) => {
                  updateField("startColumn", event.target.value);
                }}
                value={formState.startColumn}
              />
            </div>

            <div className="chart-editor__field">
              <label htmlFor="chart-row-count">Row count</label>
              <input
                id="chart-row-count"
                inputMode="numeric"
                onChange={(event) => {
                  updateField("rowCount", event.target.value);
                }}
                value={formState.rowCount}
              />
            </div>

            <div className="chart-editor__field">
              <label htmlFor="chart-column-count">Column count</label>
              <input
                id="chart-column-count"
                inputMode="numeric"
                onChange={(event) => {
                  updateField("columnCount", event.target.value);
                }}
                value={formState.columnCount}
              />
            </div>
          </section>

          {formState.chartType === "pie" ? (
            <section className="chart-editor__section chart-editor__section--grid">
              <h2 className="chart-editor__section-title">Pie dimensions</h2>

              <div className="chart-editor__field">
                <label htmlFor="chart-name-dimension">Name dimension</label>
                <input
                  id="chart-name-dimension"
                  inputMode="numeric"
                  onChange={(event) => {
                    updateField("nameDimension", event.target.value);
                  }}
                  value={formState.nameDimension}
                />
              </div>

              <div className="chart-editor__field">
                <label htmlFor="chart-value-dimension">Value dimension</label>
                <input
                  id="chart-value-dimension"
                  inputMode="numeric"
                  onChange={(event) => {
                    updateField("valueDimension", event.target.value);
                  }}
                  value={formState.valueDimension}
                />
              </div>
            </section>
          ) : (
            <section className="chart-editor__section chart-editor__section--grid">
              <h2 className="chart-editor__section-title">
                Cartesian dimensions
              </h2>

              <div className="chart-editor__field">
                <label htmlFor="chart-category-dimension">
                  X or category dimension
                </label>
                <input
                  id="chart-category-dimension"
                  inputMode="numeric"
                  onChange={(event) => {
                    updateField("categoryDimension", event.target.value);
                  }}
                  value={formState.categoryDimension}
                />
              </div>

              <div className="chart-editor__field">
                <label htmlFor="chart-value-dimensions">
                  Value dimensions
                </label>
                <input
                  id="chart-value-dimensions"
                  onChange={(event) => {
                    updateField("valueDimensions", event.target.value);
                  }}
                  placeholder="1 or 1, 2"
                  value={formState.valueDimensions}
                />
              </div>

              {formState.chartType === "line" || formState.chartType === "area" ? (
                <label className="chart-editor__checkbox">
                  <input
                    checked={formState.smooth}
                    onChange={(event) => {
                      updateField("smooth", event.target.checked);
                    }}
                    type="checkbox"
                  />
                  <span>Smooth line</span>
                </label>
              ) : null}

              {formState.chartType === "bar" || formState.chartType === "area" ? (
                <label className="chart-editor__checkbox">
                  <input
                    checked={formState.stacked}
                    onChange={(event) => {
                      updateField("stacked", event.target.checked);
                    }}
                    type="checkbox"
                  />
                  <span>Stack series</span>
                </label>
              ) : null}
            </section>
          )}

          {errorMessage ? (
            <div className="chart-editor__callout chart-editor__callout--error">
              {errorMessage}
            </div>
          ) : null}

          {validationIssues.length > 0 ? (
            <div className="chart-editor__callout">
              <strong>Fix these issues before saving</strong>
              <ul className="chart-editor__issue-list">
                {validationIssues.map((issue) => (
                  <li key={`${issue.code}:${issue.message}`}>
                    <span>{issue.code}</span>
                    <span>{issue.message}</span>
                  </li>
                ))}
              </ul>
            </div>
          ) : (
            <div className="chart-editor__callout chart-editor__callout--ok">
              This chart configuration is valid and ready to save.
            </div>
          )}
        </div>

        <footer className="chart-editor__footer">
          <button
            className="chart-editor__button chart-editor__button--secondary"
            onClick={() => {
              window.close();
            }}
            type="button"
          >
            Cancel
          </button>
          <button
            className="chart-editor__button chart-editor__button--primary"
            disabled={!canSave}
            onClick={() => {
              void handleSubmit();
            }}
            type="button"
          >
            {isSaving
              ? "Saving..."
              : request.mode === "edit"
                ? "Save chart"
                : "Create chart"}
          </button>
        </footer>
      </section>
    </main>
  );
}

function formatRangePreview(state: ChartEditorFormState): string {
  const startRow = parseTextInteger(state.startRow);
  const startColumn = parseTextInteger(state.startColumn);
  const rowCount = parseTextInteger(state.rowCount);
  const columnCount = parseTextInteger(state.columnCount);

  if (
    startRow === null ||
    startColumn === null ||
    rowCount === null ||
    columnCount === null ||
    rowCount < 1 ||
    columnCount < 1
  ) {
    return "Invalid range";
  }

  const endColumn = startColumn + columnCount - 1;
  const endRow = startRow + rowCount;

  return `${getColumnTitle(startColumn)}${startRow + 1}:${getColumnTitle(
    endColumn,
  )}${endRow}`;
}

function parseTextInteger(value: string): number | null {
  const parsed = Number.parseInt(value.trim(), 10);

  return Number.isNaN(parsed) ? null : parsed;
}
