import { type FormEvent, useEffect, useRef, useState } from "react";

import type {
  InstallerCheckUpdatesResult,
  InstallerOptions,
  InstallerStatus,
} from "./workbook-core";

interface InstallationDialogProps {
  initialMode: "check-updates" | "manage";
  onClose: () => void;
}

export function InstallationDialog({ initialMode, onClose }: InstallationDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const didAutoCheckUpdatesRef = useRef(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isBusy, setIsBusy] = useState(false);
  const [options, setOptions] = useState<InstallerOptions>({
    startMenuShortcut: false,
  });
  const [status, setStatus] = useState<InstallerStatus | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [updateResult, setUpdateResult] = useState<InstallerCheckUpdatesResult | null>(null);

  const canInstall = Boolean(status?.isPackaged && status.platform === "win32");
  const canApply = Boolean(status?.installed);
  const canCheckUpdates = Boolean(status?.canManageInstalledInstance);

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
      .getInstallerStatus()
      .then((nextStatus) => {
        if (!isMounted) {
          return;
        }

        setStatus(nextStatus);
        setOptions(nextStatus.options);
      })
      .catch((error) => {
        if (isMounted) {
          setErrorMessage(getErrorMessage(error, "Installation status could not be loaded."));
        }
      });

    return () => {
      isMounted = false;
    };
  }, []);

  useEffect(() => {
    if (initialMode !== "check-updates" || !status || didAutoCheckUpdatesRef.current) {
      return;
    }

    didAutoCheckUpdatesRef.current = true;
    void checkUpdates(false);
  }, [initialMode, status]);

  const updateOption = (option: keyof InstallerOptions, value: boolean) => {
    setOptions((current) => ({
      ...current,
      [option]: value,
    }));
  };

  const runOperation = async (
    operation: () => Promise<{ message: string; status: InstallerStatus }>,
  ) => {
    setIsBusy(true);
    setErrorMessage(null);
    setSuccessMessage(null);

    try {
      const result = await operation();

      setStatus(result.status);
      setOptions(result.status.options);
      setSuccessMessage(result.message);
    } catch (error) {
      setErrorMessage(getErrorMessage(error, "Installation operation failed."));
    } finally {
      setIsBusy(false);
    }
  };

  const checkUpdates = async (startUpdate: boolean) => {
    setIsBusy(true);
    setErrorMessage(null);
    setSuccessMessage(null);

    try {
      const result = await window.appShell.checkForInstallerUpdates({
        restart: true,
        startUpdate,
      });

      setStatus(result.status);
      setUpdateResult(result);

      if (result.updateStarted || !result.updateAvailable) {
        setSuccessMessage(result.message);
      }
    } catch (error) {
      setErrorMessage(getErrorMessage(error, "Update check failed."));
    } finally {
      setIsBusy(false);
    }
  };

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!status?.installed) {
      void runOperation(() => window.appShell.installCurrentApp(options));
      return;
    }

    void runOperation(() => window.appShell.applyInstallerOptions(options));
  };

  return (
    <dialog
      aria-labelledby="installation-title"
      className="chart-editor-dialog"
      onCancel={(event) => {
        event.preventDefault();
        onClose();
      }}
      ref={dialogRef}
    >
      <main className="chart-editor-window installation-window">
        <form className="chart-editor-panel installation-panel" onSubmit={handleSubmit}>
          <header className="chart-editor__header">
            <p className="chart-editor__eyebrow" id="installation-title">
              Installation
            </p>
            <h1 className="chart-editor__title">Spready</h1>
            <p className="chart-editor__subtitle">
              {status?.installed
                ? "Installed for this Windows user."
                : "Install for this Windows user."}
            </p>
          </header>

          <div className="chart-editor__body">
            {status ? (
              <section className="chart-editor__section installation__section">
                <div className="installation__detail">
                  <span>Version</span>
                  <strong>{status.currentVersion}</strong>
                </div>
                <div className="installation__detail">
                  <span>Install folder</span>
                  <strong>{status.installDirectory}</strong>
                </div>
                <label className="chart-editor__checkbox">
                  <input
                    checked={options.startMenuShortcut}
                    disabled={!canInstall || isBusy}
                    onChange={(event) => {
                      updateOption("startMenuShortcut", event.target.checked);
                    }}
                    type="checkbox"
                  />
                  Add Start Menu shortcut
                </label>
              </section>
            ) : (
              <div className="chart-editor__loading">Loading installation status...</div>
            )}

            {status && !canInstall ? (
              <div className="chart-editor__callout">
                Installation and updates are available from a packaged Windows build.
              </div>
            ) : null}

            {status?.installed && !status.canManageInstalledInstance ? (
              <div className="chart-editor__callout">
                Updates and uninstall are available when running the installed Spready executable.
              </div>
            ) : null}

            {updateResult?.updateAvailable ? (
              <div className="chart-editor__callout chart-editor__callout--ok">
                {updateResult.message}
              </div>
            ) : null}

            {successMessage ? (
              <div className="chart-editor__callout chart-editor__callout--ok">
                {successMessage}
              </div>
            ) : null}

            {errorMessage ? (
              <div className="chart-editor__callout chart-editor__callout--error">
                {errorMessage}
              </div>
            ) : null}
          </div>

          <footer className="chart-editor__footer">
            <span className="chart-editor__status" role="status">
              {status?.installed ? "Installed" : "Not installed"}
            </span>
            <div className="chart-editor__actions">
              <button
                className="chart-editor__button chart-editor__button--secondary"
                disabled={isBusy || !canCheckUpdates}
                onClick={() => {
                  void checkUpdates(false);
                }}
                type="button"
              >
                Check Updates
              </button>
              {updateResult?.updateAvailable ? (
                <button
                  className="chart-editor__button chart-editor__button--primary"
                  disabled={isBusy}
                  onClick={() => {
                    void checkUpdates(true);
                  }}
                  type="button"
                >
                  Update
                </button>
              ) : null}
              <button
                className="chart-editor__button chart-editor__button--secondary"
                disabled={isBusy || !status?.canManageInstalledInstance}
                onClick={() => {
                  void runOperation(() => window.appShell.startUninstall());
                }}
                type="button"
              >
                Uninstall
              </button>
              <button
                className="chart-editor__button chart-editor__button--secondary"
                onClick={onClose}
                type="button"
              >
                Close
              </button>
              <button
                className="chart-editor__button chart-editor__button--primary"
                disabled={isBusy || !canInstall || (status?.installed && !canApply)}
                type="submit"
              >
                {status?.installed ? "Apply" : "Install"}
              </button>
            </div>
          </footer>
        </form>
      </main>
    </dialog>
  );
}

function getErrorMessage(error: unknown, fallback: string) {
  return error instanceof Error ? error.message : fallback;
}
