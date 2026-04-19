import { useEffect } from "react";

import type { ToastNotification } from "./toast-state";

type ToastViewportProps = {
  onDismiss: (toastId: string) => void;
  toasts: readonly ToastNotification[];
};

function capitalizeToastKind(kind: ToastNotification["kind"]): string {
  return `${kind.slice(0, 1).toUpperCase()}${kind.slice(1)}`;
}

function ToastAutoDismiss({
  dismissAfterMs,
  onDismiss,
  toastId,
}: {
  dismissAfterMs: number;
  onDismiss: (toastId: string) => void;
  toastId: string;
}) {
  useEffect(() => {
    if (dismissAfterMs <= 0) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      onDismiss(toastId);
    }, dismissAfterMs);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [dismissAfterMs, onDismiss, toastId]);

  return null;
}

export function ToastViewport({ onDismiss, toasts }: ToastViewportProps) {
  const visibleToasts = [...toasts].reverse();

  return (
    <div aria-atomic="false" aria-live="polite" className="toast-viewport">
      {visibleToasts.map((toast) => (
        <article
          className={`toast toast--${toast.kind}`}
          key={toast.id}
          role={toast.kind === "error" ? "alert" : "status"}
        >
          <ToastAutoDismiss
            dismissAfterMs={toast.dismissAfterMs}
            onDismiss={onDismiss}
            toastId={toast.id}
          />
          <div className="toast__header">
            <div className="toast__content">
              <div className="toast__label-row">
                <span className="toast__label">
                  {capitalizeToastKind(toast.kind)}
                </span>
                {toast.occurrenceCount > 1 ? (
                  <span
                    className="toast__count"
                    aria-label={`Repeated ${toast.occurrenceCount} times`}
                  >
                    x{toast.occurrenceCount}
                  </span>
                ) : null}
              </div>
              <div className="toast__title">{toast.title}</div>
              {toast.description ? (
                <p className="toast__description">{toast.description}</p>
              ) : null}
            </div>

            <button
              aria-label={`Dismiss ${toast.kind} notification`}
              className="toast__dismiss"
              onClick={() => {
                onDismiss(toast.id);
              }}
              type="button"
            >
              Close
            </button>
          </div>
        </article>
      ))}
    </div>
  );
}
