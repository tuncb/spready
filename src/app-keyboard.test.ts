import assert from "node:assert/strict";
import { test } from "node:test";

import { shouldCloseWorkbookSearchOnEscape } from "./app-keyboard";

function createKeyEvent(
  overrides: Partial<
    Pick<KeyboardEvent, "altKey" | "ctrlKey" | "key" | "metaKey" | "shiftKey" | "target">
  > = {},
) {
  return {
    altKey: false,
    ctrlKey: false,
    key: "Escape",
    metaKey: false,
    shiftKey: false,
    target: null,
    ...overrides,
  };
}

test("shouldCloseWorkbookSearchOnEscape closes active search from the grid", () => {
  assert.equal(
    shouldCloseWorkbookSearchOnEscape(createKeyEvent(), {
      hasSearchQuery: true,
      isFormulaInputFocused: false,
      isSearchOpen: true,
    }),
    true,
  );
});

test("shouldCloseWorkbookSearchOnEscape ignores inactive search states", () => {
  assert.equal(
    shouldCloseWorkbookSearchOnEscape(createKeyEvent(), {
      hasSearchQuery: false,
      isFormulaInputFocused: false,
      isSearchOpen: true,
    }),
    false,
  );
  assert.equal(
    shouldCloseWorkbookSearchOnEscape(createKeyEvent(), {
      hasSearchQuery: true,
      isFormulaInputFocused: false,
      isSearchOpen: false,
    }),
    false,
  );
});

test("shouldCloseWorkbookSearchOnEscape does not steal modified or non-escape keys", () => {
  assert.equal(
    shouldCloseWorkbookSearchOnEscape(createKeyEvent({ ctrlKey: true }), {
      hasSearchQuery: true,
      isFormulaInputFocused: false,
      isSearchOpen: true,
    }),
    false,
  );
  assert.equal(
    shouldCloseWorkbookSearchOnEscape(createKeyEvent({ key: "Enter" }), {
      hasSearchQuery: true,
      isFormulaInputFocused: false,
      isSearchOpen: true,
    }),
    false,
  );
});

test("shouldCloseWorkbookSearchOnEscape leaves formula editing alone", () => {
  assert.equal(
    shouldCloseWorkbookSearchOnEscape(createKeyEvent(), {
      hasSearchQuery: true,
      isFormulaInputFocused: true,
      isSearchOpen: true,
    }),
    false,
  );
});
