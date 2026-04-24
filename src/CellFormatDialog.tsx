import { useEffect, useRef, useState } from "react";

import type {
  SheetRangeRequest,
  WorkbookCellStyle,
  WorkbookTransactionOperation,
} from "./workbook-core";

interface CellFormatDialogProps {
  expectedVersion: number;
  initialStyle?: WorkbookCellStyle;
  onClose: () => void;
  onVersionConflict: (message: string) => void;
  ranges: SheetRangeRequest[];
}

const DEFAULT_TEXT_COLOR = "#0f172a";
const DEFAULT_FILL_COLOR = "#ffffff";

export function CellFormatDialog({
  expectedVersion,
  initialStyle,
  onClose,
  onVersionConflict,
  ranges,
}: CellFormatDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [formState, setFormState] = useState<WorkbookCellStyle>(() => initialStyle ?? {});
  const [isSaving, setIsSaving] = useState(false);

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

  const updateStyle = (patch: WorkbookCellStyle) => {
    setFormState((current) => ({
      ...current,
      ...patch,
    }));
  };

  const handleSubmit = async () => {
    setIsSaving(true);
    setErrorMessage(null);

    try {
      const style = normalizeDialogCellStyle(formState);
      const operations: WorkbookTransactionOperation[] = ranges.map((range) => ({
        ...range,
        style,
        type: "setRangeStyle",
      }));

      await window.appShell.applyTransaction({
        expectedVersion,
        operations,
      });
      onClose();
    } catch (error) {
      const message = error instanceof Error ? error.message : "Cell format could not be saved.";

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
      aria-labelledby="cell-format-title"
      className="chart-editor-dialog"
      onCancel={(event) => {
        event.preventDefault();
        onClose();
      }}
      ref={dialogRef}
    >
      <main className="chart-editor-window">
        <section className="chart-editor-panel">
          <header className="chart-editor__header">
            <p className="chart-editor__eyebrow" id="cell-format-title">
              Format cells
            </p>
          </header>

          <div className="chart-editor__body">
            <section className="chart-editor__section cell-format__section">
              <div className="cell-format__toggle-group">
                <button
                  className={
                    formState.bold ? "cell-format__toggle is-active" : "cell-format__toggle"
                  }
                  onClick={() => {
                    updateStyle({ bold: !formState.bold });
                  }}
                  type="button"
                >
                  B
                </button>
                <button
                  className={
                    formState.italic
                      ? "cell-format__toggle cell-format__toggle--italic is-active"
                      : "cell-format__toggle cell-format__toggle--italic"
                  }
                  onClick={() => {
                    updateStyle({ italic: !formState.italic });
                  }}
                  type="button"
                >
                  I
                </button>
              </div>

              <div className="chart-editor__field">
                <label htmlFor="cell-format-family">Font family</label>
                <select
                  id="cell-format-family"
                  onChange={(event) => {
                    updateStyle({
                      fontFamily: event.target.value || undefined,
                    });
                  }}
                  value={formState.fontFamily ?? ""}
                >
                  <option value="">Aptos</option>
                  <option value="Arial">Arial</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Consolas">Consolas</option>
                  <option value="Georgia">Georgia</option>
                  <option value="Times New Roman">Times</option>
                </select>
              </div>

              <div className="chart-editor__field">
                <label htmlFor="cell-format-size">Font size</label>
                <select
                  id="cell-format-size"
                  onChange={(event) => {
                    updateStyle({
                      fontSize:
                        event.target.value === ""
                          ? undefined
                          : Number.parseInt(event.target.value, 10),
                    });
                  }}
                  value={formState.fontSize?.toString() ?? ""}
                >
                  <option value="">13</option>
                  {[10, 11, 12, 14, 16, 18, 20, 24, 28, 32].map((fontSize) => (
                    <option key={fontSize} value={fontSize}>
                      {fontSize}
                    </option>
                  ))}
                </select>
              </div>

              <div className="chart-editor__field">
                <label htmlFor="cell-format-align">Alignment</label>
                <select
                  id="cell-format-align"
                  onChange={(event) => {
                    updateStyle({
                      horizontalAlign:
                        event.target.value === ""
                          ? undefined
                          : (event.target.value as WorkbookCellStyle["horizontalAlign"]),
                    });
                  }}
                  value={formState.horizontalAlign ?? ""}
                >
                  <option value="">Left</option>
                  <option value="center">Center</option>
                  <option value="right">Right</option>
                </select>
              </div>

              <div className="chart-editor__field">
                <label htmlFor="cell-format-text-color">Text color</label>
                <input
                  id="cell-format-text-color"
                  onChange={(event) => {
                    updateStyle({
                      textColor: event.target.value,
                    });
                  }}
                  type="color"
                  value={getColorInputValue(formState.textColor, DEFAULT_TEXT_COLOR)}
                />
              </div>

              <div className="chart-editor__field">
                <label htmlFor="cell-format-fill-color">Fill color</label>
                <input
                  id="cell-format-fill-color"
                  onChange={(event) => {
                    updateStyle({
                      backgroundColor: event.target.value,
                    });
                  }}
                  type="color"
                  value={getColorInputValue(formState.backgroundColor, DEFAULT_FILL_COLOR)}
                />
              </div>

              <label className="chart-editor__checkbox cell-format__checkbox">
                <input
                  checked={formState.wrapText ?? false}
                  onChange={(event) => {
                    updateStyle({
                      wrapText: event.target.checked,
                    });
                  }}
                  type="checkbox"
                />
                <span>Wrap text</span>
              </label>
            </section>

            {errorMessage ? (
              <div className="chart-editor__callout chart-editor__callout--error">
                {errorMessage}
              </div>
            ) : null}
          </div>

          <footer className="chart-editor__footer">
            <span className="chart-editor__status" role="status">
              {formatRangeCount(ranges)}
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
                disabled={isSaving || ranges.length === 0}
                onClick={() => {
                  void handleSubmit();
                }}
                type="button"
              >
                {isSaving ? "Saving..." : "Apply"}
              </button>
            </div>
          </footer>
        </section>
      </main>
    </dialog>
  );
}

function normalizeDialogCellStyle(style: WorkbookCellStyle): WorkbookCellStyle | undefined {
  const normalized: WorkbookCellStyle = {};

  if (style.backgroundColor) {
    normalized.backgroundColor = style.backgroundColor;
  }

  if (style.bold) {
    normalized.bold = true;
  }

  if (style.fontFamily) {
    normalized.fontFamily = style.fontFamily;
  }

  if (style.fontSize) {
    normalized.fontSize = style.fontSize;
  }

  if (style.horizontalAlign) {
    normalized.horizontalAlign = style.horizontalAlign;
  }

  if (style.italic) {
    normalized.italic = true;
  }

  if (style.textColor) {
    normalized.textColor = style.textColor;
  }

  if (style.wrapText) {
    normalized.wrapText = true;
  }

  return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function getColorInputValue(value: string | undefined, fallback: string) {
  return /^#[0-9a-f]{6}$/i.test(value ?? "") ? (value as string) : fallback;
}

function formatRangeCount(ranges: readonly SheetRangeRequest[]) {
  const cellCount = ranges.reduce((total, range) => total + range.rowCount * range.columnCount, 0);

  return cellCount === 1 ? "1 cell selected" : `${cellCount} cells selected`;
}

function isExpectedVersionConflict(message: string): boolean {
  return message.startsWith("Expected workbook version ");
}
