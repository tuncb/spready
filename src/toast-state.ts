export type ToastKind = "error" | "info" | "success" | "warning";

export interface ToastInput {
  description?: string;
  dismissAfterMs?: number;
  kind: ToastKind;
  title: string;
}

export interface ToastNotification {
  description?: string;
  dismissAfterMs: number;
  id: string;
  kind: ToastKind;
  occurrenceCount: number;
  signature: string;
  title: string;
}

const DEFAULT_TOAST_LIMIT = 4;
const DEFAULT_DISMISS_AFTER_MS: Record<ToastKind, number> = {
  error: 10_000,
  info: 5_000,
  success: 4_000,
  warning: 7_000,
};

let toastSequence = 0;

function normalizeToastText(value: string | undefined): string | undefined {
  const trimmed = value?.trim();

  return trimmed && trimmed.length > 0 ? trimmed : undefined;
}

function createToastSignature(input: ToastInput): string {
  return [
    input.kind,
    normalizeToastText(input.title) ?? "Notification",
    normalizeToastText(input.description) ?? "",
  ].join("\u0000");
}

export function enqueueToast(
  current: readonly ToastNotification[],
  input: ToastInput,
  options?: {
    limit?: number;
    now?: number;
  },
): ToastNotification[] {
  const limit = options?.limit ?? DEFAULT_TOAST_LIMIT;
  const now = options?.now ?? Date.now();
  const title = normalizeToastText(input.title) ?? "Notification";
  const description = normalizeToastText(input.description);
  const signature = createToastSignature({
    ...input,
    description,
    title,
  });
  const existingToast = current.find((toast) => toast.signature === signature);
  const nextToast: ToastNotification = {
    description,
    dismissAfterMs: input.dismissAfterMs ?? DEFAULT_DISMISS_AFTER_MS[input.kind],
    id: `toast-${now}-${toastSequence}`,
    kind: input.kind,
    occurrenceCount: (existingToast?.occurrenceCount ?? 0) + 1,
    signature,
    title,
  };

  toastSequence += 1;

  const nextQueue = current.filter((toast) => toast.signature !== signature).concat(nextToast);

  return limit > 0 ? nextQueue.slice(-limit) : nextQueue;
}

export function removeToast(
  current: readonly ToastNotification[],
  toastId: string,
): ToastNotification[] {
  return current.filter((toast) => toast.id !== toastId);
}
