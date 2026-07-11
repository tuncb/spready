import { useEffect, useRef, useState } from "react";

import { isDialogBackdropClick } from "./dialog-events";
import type { ControlAppInfo } from "./workbook-core";

interface AboutDialogProps {
  onClose: () => void;
}

export function AboutDialog({ onClose }: AboutDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [appInfo, setAppInfo] = useState<ControlAppInfo | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

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
    let isMounted = true;

    void window.appShell
      .getAppInfo()
      .then((info) => {
        if (isMounted) {
          setAppInfo(info);
        }
      })
      .catch((error) => {
        if (isMounted) {
          setErrorMessage(
            error instanceof Error ? error.message : "App details could not be loaded.",
          );
        }
      });

    return () => {
      isMounted = false;
    };
  }, []);

  const controlAddress = appInfo
    ? `tcp://${appInfo.controlServer.host}:${appInfo.controlServer.port}`
    : "Loading...";

  return (
    <dialog
      aria-labelledby="about-title"
      className="chart-editor-dialog"
      onCancel={(event) => {
        event.preventDefault();
        onClose();
      }}
      onClick={(event) => {
        if (isDialogBackdropClick(event)) {
          onClose();
        }
      }}
      ref={dialogRef}
    >
      <main className="chart-editor-window about-window">
        <section className="chart-editor-panel about-panel">
          <header className="chart-editor__header">
            <p className="chart-editor__eyebrow">About</p>
            <h1 className="chart-editor__title" id="about-title">
              {appInfo?.name ?? window.appShell.name}
            </h1>
            <p className="chart-editor__subtitle">Desktop spreadsheet application</p>
          </header>

          <div className="chart-editor__body">
            <section className="chart-editor__section about__section">
              <div className="about__detail">
                <span>Version</span>
                <strong>{appInfo?.version ?? "Loading..."}</strong>
              </div>
              <div className="about__detail">
                <span>Automation</span>
                <strong>{controlAddress}</strong>
              </div>
            </section>

            {errorMessage ? (
              <div className="chart-editor__callout chart-editor__callout--error">
                {errorMessage}
              </div>
            ) : null}
          </div>

          <footer className="chart-editor__footer about__footer">
            <button
              autoFocus
              className="chart-editor__button chart-editor__button--primary"
              onClick={onClose}
              type="button"
            >
              Close
            </button>
          </footer>
        </section>
      </main>
    </dialog>
  );
}
