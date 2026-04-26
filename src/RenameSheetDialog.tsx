import { type FormEvent, useEffect, useRef, useState } from "react";

import type { WorkbookSummary } from "./workbook-core";

interface RenameSheetDialogProps {
  expectedVersion: number;
  initialName: string;
  onClose: () => void;
  onSaved: (summary: WorkbookSummary) => void;
  onVersionConflict: (message: string) => void;
  sheetId: string;
}

export function RenameSheetDialog({
  expectedVersion,
  initialName,
  onClose,
  onSaved,
  onVersionConflict,
  sheetId,
}: RenameSheetDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [name, setName] = useState(initialName);
  const trimmedName = name.trim();
  const canSave = trimmedName.length > 0 && trimmedName !== initialName.trim() && !isSaving;

  useEffect(() => {
    const dialog = dialogRef.current;

    if (!dialog || dialog.open) {
      return;
    }

    dialog.showModal();
    inputRef.current?.select();

    return () => {
      if (dialog.open) {
        dialog.close();
      }
    };
  }, []);

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!canSave) {
      return;
    }

    setIsSaving(true);
    setErrorMessage(null);

    try {
      const result = await window.appShell.applyTransaction({
        expectedVersion,
        operations: [
          {
            name: trimmedName,
            sheetId,
            type: "renameSheet",
          },
        ],
      });

      onSaved(result.summary);
      onClose();
    } catch (error) {
      const message = error instanceof Error ? error.message : "Sheet could not be renamed.";

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
      aria-labelledby="rename-sheet-title"
      className="chart-editor-dialog"
      onCancel={(event) => {
        event.preventDefault();
        onClose();
      }}
      ref={dialogRef}
    >
      <main className="chart-editor-window rename-sheet-window">
        <form className="chart-editor-panel rename-sheet-panel" onSubmit={handleSubmit}>
          <header className="chart-editor__header">
            <p className="chart-editor__eyebrow" id="rename-sheet-title">
              Rename sheet
            </p>
            <h1 className="chart-editor__title">Sheet name</h1>
            <p className="chart-editor__subtitle">Change the active sheet name.</p>
          </header>

          <div className="chart-editor__body">
            <section className="chart-editor__section rename-sheet__section">
              <div className="chart-editor__field">
                <label htmlFor="rename-sheet-name">Name</label>
                <input
                  id="rename-sheet-name"
                  onChange={(event) => {
                    setName(event.target.value);
                  }}
                  ref={inputRef}
                  value={name}
                />
              </div>
            </section>

            {errorMessage ? (
              <div className="chart-editor__callout chart-editor__callout--error">
                {errorMessage}
              </div>
            ) : null}
          </div>

          <footer className="chart-editor__footer">
            <span className="chart-editor__status" role="status">
              {trimmedName.length === 0 ? "Name is required" : initialName}
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
                type="submit"
              >
                {isSaving ? "Saving..." : "Rename"}
              </button>
            </div>
          </footer>
        </form>
      </main>
    </dialog>
  );
}

function isExpectedVersionConflict(message: string): boolean {
  return message.startsWith("Expected workbook version ");
}
