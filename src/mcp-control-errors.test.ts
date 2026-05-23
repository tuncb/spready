import assert from "node:assert/strict";
import { test } from "node:test";

import { formatControlConnectionError } from "./mcp-control-errors";

test("formatControlConnectionError explains refused control connections", () => {
  const error = Object.assign(new Error("connect ECONNREFUSED 127.0.0.1:45731"), {
    code: "ECONNREFUSED",
  });

  assert.equal(
    formatControlConnectionError(
      {
        host: "127.0.0.1",
        port: 45731,
        source: "default",
      },
      error,
    ),
    "Could not connect to the Spready TCP control server at tcp://127.0.0.1:45731. Open Spready at this address through MCP or manually.",
  );
});

test("formatControlConnectionError preserves unexpected connection details", () => {
  assert.equal(
    formatControlConnectionError(
      {
        host: "127.0.0.1",
        port: 45731,
        source: "default",
      },
      new Error("TLS handshake failed"),
    ),
    "Could not connect to the Spready TCP control server at tcp://127.0.0.1:45731. TLS handshake failed",
  );
});
