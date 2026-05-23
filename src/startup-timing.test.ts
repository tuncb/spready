import assert from "node:assert/strict";
import { test } from "node:test";

import { formatStartupTimingLog, StartupTimer } from "./startup-timing";

test("formatStartupTimingLog includes scope, elapsed, total, event, and detail", () => {
  assert.equal(
    formatStartupTimingLog("spready-mcp", "launch-app-start", 12.4, 98.6, "pid=123"),
    "[spready-mcp] +12ms total=99ms launch-app-start pid=123",
  );
});

test("StartupTimer reports elapsed and total milliseconds through the sink", () => {
  const messages: string[] = [];
  const times = [1000, 1010, 1045];
  const timer = new StartupTimer(
    "spready-main",
    (message) => {
      messages.push(message);
    },
    () => times.shift() ?? 1045,
  );

  timer.log("app-when-ready");
  timer.log("control-server-started", "tcp://127.0.0.1:45731");

  assert.deepEqual(messages, [
    "[spready-main] +10ms total=10ms app-when-ready",
    "[spready-main] +35ms total=45ms control-server-started tcp://127.0.0.1:45731",
  ]);
});
