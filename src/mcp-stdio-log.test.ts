import assert from "node:assert/strict";
import { test } from "node:test";

import {
  formatConnectedStartupLog,
  formatDisconnectedStartupLog,
  formatReadyStartupLog,
} from "./mcp-stdio-log";

test("formatDisconnectedStartupLog separates startup details across lines", () => {
  assert.equal(
    formatDisconnectedStartupLog(
      "C:\\Users\\tuncb\\AppData\\Local\\Temp\\spready-control.json",
      "Could not connect to the Spready TCP control server at tcp://127.0.0.1:45731.",
    ),
    [
      "Spready MCP stdio wrapper started without a control connection.",
      "Action: call open_spready_app to connect or launch Spready.",
      "Discovery file: C:\\Users\\tuncb\\AppData\\Local\\Temp\\spready-control.json",
      "Detail: Could not connect to the Spready TCP control server at tcp://127.0.0.1:45731.",
    ].join("\n"),
  );
});

test("formatConnectedStartupLog separates connected target details across lines", () => {
  assert.equal(
    formatConnectedStartupLog({
      host: "127.0.0.1",
      port: 45731,
      source: "default",
    }),
    [
      "Spready MCP stdio wrapper connected.",
      "Address: tcp://127.0.0.1:45731",
      "Source: default",
    ].join("\n"),
  );
});

test("formatReadyStartupLog separates ready action across lines", () => {
  assert.equal(
    formatReadyStartupLog(),
    ["Spready MCP stdio wrapper ready.", "Action: call open_spready_app to connect."].join("\n"),
  );
});
