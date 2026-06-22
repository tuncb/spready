import assert from "node:assert/strict";
import { test } from "node:test";

import { getMainHelpText, parseMainStartupOptions } from "./app-startup";

test("parseMainStartupOptions reads console output file paths", () => {
  assert.deepEqual(parseMainStartupOptions(["--console-output", "C:\\workbooks\\budget.spready"]), {
    consoleOutputFilePath: "C:\\workbooks\\budget.spready",
    help: false,
  });
  assert.deepEqual(parseMainStartupOptions(["--console-output=/tmp/budget.spready"]), {
    consoleOutputFilePath: "/tmp/budget.spready",
    help: false,
  });
});

test("parseMainStartupOptions recognizes help flags", () => {
  assert.deepEqual(parseMainStartupOptions(["--help"]), {
    help: true,
  });
  assert.deepEqual(parseMainStartupOptions(["-h"]), {
    help: true,
  });
});

test("parseMainStartupOptions rejects missing console output values", () => {
  assert.throws(() => parseMainStartupOptions(["--console-output"]), /Missing value/);
  assert.throws(() => parseMainStartupOptions(["--console-output="]), /Missing value/);
});

test("getMainHelpText documents console output mode", () => {
  const helpText = getMainHelpText("spready");

  assert.match(helpText, /Usage: spready \[options\]/);
  assert.match(helpText, /--console-output FILE/);
});
