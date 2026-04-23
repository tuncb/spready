import { useEffect, useRef, useState } from "react";

import {
  buildChartEditorOperations,
  createChartEditorFormState,
  getChartEditorSheetId,
  getChartEditorValidationIssues,
  type ChartEditorFormState,
  type ChartEditorWindowRequest,
} from "./chart-editor-state";
import { type WorkbookChart, type WorkbookSummary } from "./workbook-core";

interface ChartEditorDialogProps {
  expectedVersion: number;
  onClose: () => void;
  onVersionConflict: (message: string) => void;
  request: ChartEditorWindowRequest;
}

export function ChartEditorDialog({
  expectedVersion,
  onClose,
  onVersionConflict,
  request,
}: ChartEditorDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [chart, setChart] = useState<WorkbookChart | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [formState, setFormState] = useState<ChartEditorFormState | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [summary, setSummary] = useState<WorkbookSummary | null>(null);

  useEffect(() => {
    const dialog = dialogRef.current;

    if (!dialog || dialog.open) {
      return;
    }

    dialog.showModal();

    return () => {
      if (dialog.open) {
        dialog.close();
      }
    };
  }, []);

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
  const validationIssues =
    summary && formState && sheetId
      ? getChartEditorValidationIssues(
          formState,
          sheetId,
          summary,
          request.mode === "edit" ? request.chartId : undefined,
        )
      : [];
  const validationStatus =
    validationIssues.length === 0
      ? "Valid configuration"
      : "Fix issues before saving";
  const canSave =
    !isLoading &&
    !isSaving &&
    summary !== null &&
    formState !== null &&
    sheetId !== null &&
    validationIssues.length === 0;

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

  const handleChartTypeChange = (
    nextChartType: ChartEditorFormState["chartType"],
  ) => {
    setFormState((current) => {
      if (current === null) {
        return current;
      }

      return {
        ...current,
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
      };
    });
  };

  const handleSubmit = async () => {
    if (!formState || !sheetId) {
      return;
    }

    setIsSaving(true);
    setErrorMessage(null);

    try {
      await window.appShell.applyTransaction({
        expectedVersion,
        operations: buildChartEditorOperations(request, sheetId, formState),
      });
      onClose();
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Chart could not be saved.";

      if (isExpectedVersionConflict(message)) {
        onVersionConflict(message);
        return;
      }

      setErrorMessage(message);
      setIsSaving(false);
    }
  };

  return (
    <dialog
      aria-labelledby="chart-editor-title"
      className="chart-editor-dialog"
      onCancel={(event) => {
        event.preventDefault();
        onClose();
      }}
      ref={dialogRef}
    >
      {isLoading ||
      formState === null ||
      summary === null ||
      sheetId === null ? (
        <main className="chart-editor-window">
          <section className="chart-editor-panel">
            <div className="chart-editor__loading">
              {errorMessage ?? "Loading chart editor..."}
            </div>
            <footer className="chart-editor__footer">
              <button
                className="chart-editor__button chart-editor__button--secondary"
                onClick={onClose}
                type="button"
              >
                Close
              </button>
            </footer>
          </section>
        </main>
      ) : (
        <main className="chart-editor-window">
          <section className="chart-editor-panel">
            <header className="chart-editor__header">
              <div>
                <p className="chart-editor__eyebrow">
                  {request.mode === "edit" ? "Edit chart" : "Insert chart"}
                </p>
                <h1 className="chart-editor__title" id="chart-editor-title">
                  {request.mode === "edit" ? "Edit Chart" : "Create Chart"}
                </h1>
                <p className="chart-editor__subtitle">
                  {formState.sourceRange}
                </p>
              </div>
            </header>

            <div className="chart-editor__body">
              <section className="chart-editor__section chart-editor__section--basics">
                <div className="chart-editor__field chart-editor__field--name">
                  <label htmlFor="chart-name">Chart name</label>
                  <input
                    id="chart-name"
                    onChange={(event) => {
                      updateField("name", event.target.value);
                    }}
                    value={formState.name}
                  />
                </div>

                <div className="chart-editor__field chart-editor__field--type">
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

              <section className="chart-editor__section chart-editor__section--data">
                <div className="chart-editor__field chart-editor__field--range">
                  <label htmlFor="chart-source-range">Data range</label>
                  <input
                    autoCapitalize="characters"
                    id="chart-source-range"
                    onChange={(event) => {
                      updateField("sourceRange", event.target.value);
                    }}
                    spellCheck={false}
                    value={formState.sourceRange}
                  />
                </div>

                <div className="chart-editor__field chart-editor__field--layout">
                  <label htmlFor="chart-layout">Series layout</label>
                  <select
                    id="chart-layout"
                    onChange={(event) => {
                      updateField(
                        "seriesLayoutBy",
                        event.target
                          .value as ChartEditorFormState["seriesLayoutBy"],
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

              {formState.chartType === "pie" ? (
                <section className="chart-editor__section chart-editor__section--grid">
                  <h2 className="chart-editor__section-title">
                    Pie dimensions
                  </h2>

                  <div className="chart-editor__field chart-editor__field--short">
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

                  <div className="chart-editor__field chart-editor__field--short">
                    <label htmlFor="chart-value-dimension">
                      Value dimension
                    </label>
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

                  <div className="chart-editor__field chart-editor__field--short">
                    <label htmlFor="chart-category-dimension">
                      X axis / category column
                    </label>
                    <input
                      id="chart-category-dimension"
                      inputMode="numeric"
                      onChange={(event) => {
                        updateField("categoryDimension", event.target.value);
                      }}
                      placeholder="0"
                      value={formState.categoryDimension}
                    />
                  </div>

                  <div className="chart-editor__field chart-editor__field--medium">
                    <label htmlFor="chart-value-dimensions">
                      Y value columns
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

                  {formState.chartType === "line" ||
                  formState.chartType === "area" ? (
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

                  {formState.chartType === "bar" ||
                  formState.chartType === "area" ? (
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
              ) : null}
            </div>

            <footer className="chart-editor__footer">
              <span
                className={
                  validationIssues.length === 0
                    ? "chart-editor__status chart-editor__status--ok"
                    : "chart-editor__status"
                }
              >
                {validationStatus}
              </span>
              <div className="chart-editor__actions">
                <button
                  className="chart-editor__button chart-editor__button--secondary"
                  onClick={onClose}
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
                      ? "Save Chart"
                      : "Create Chart"}
                </button>
              </div>
            </footer>
          </section>
        </main>
      )}
    </dialog>
  );
}

function isExpectedVersionConflict(message: string): boolean {
  return message.startsWith("Expected workbook version ");
}
