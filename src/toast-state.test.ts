import assert from "node:assert/strict";
import { test } from "node:test";

import { enqueueToast, removeToast } from "./toast-state";

test("enqueueToast adds a normalized toast with default dismissal timing", () => {
  const queue = enqueueToast([], {
    kind: "error",
    title: "  Save failed  ",
  });

  assert.equal(queue.length, 1);
  assert.equal(queue[0]?.kind, "error");
  assert.equal(queue[0]?.title, "Save failed");
  assert.equal(queue[0]?.dismissAfterMs, 10_000);
  assert.equal(queue[0]?.occurrenceCount, 1);
});

test("enqueueToast deduplicates matching toasts and increments the count", () => {
  const firstQueue = enqueueToast(
    [],
    {
      description: "Disk full",
      kind: "error",
      title: "Export failed",
    },
    { now: 100 },
  );
  const secondQueue = enqueueToast(
    firstQueue,
    {
      description: "Disk full",
      kind: "error",
      title: "Export failed",
    },
    { now: 200 },
  );

  assert.equal(secondQueue.length, 1);
  assert.equal(secondQueue[0]?.occurrenceCount, 2);
  assert.notEqual(secondQueue[0]?.id, firstQueue[0]?.id);
});

test("enqueueToast keeps only the newest toasts within the queue limit", () => {
  const queue = [
    { kind: "info", title: "One" },
    { kind: "info", title: "Two" },
    { kind: "info", title: "Three" },
  ] as const;
  const result = queue.reduce(
    (current, toast) => enqueueToast(current, toast, { limit: 2 }),
    [] as ReturnType<typeof enqueueToast>,
  );

  assert.deepEqual(
    result.map((toast) => toast.title),
    ["Two", "Three"],
  );
});

test("removeToast deletes a toast by id", () => {
  const queue = enqueueToast([], {
    kind: "warning",
    title: "Heads up",
  });
  const toastId = queue[0]?.id;

  assert.ok(toastId);
  assert.deepEqual(removeToast(queue, toastId), []);
});
