import assert from "node:assert/strict";
import { test } from "node:test";

import { isDialogBackdropClick } from "./dialog-events";

test("isDialogBackdropClick identifies clicks targeting the dialog backdrop", () => {
  const dialog = new EventTarget();

  assert.equal(isDialogBackdropClick({ currentTarget: dialog, target: dialog }), true);
});

test("isDialogBackdropClick ignores clicks originating inside the dialog", () => {
  const dialog = new EventTarget();
  const content = new EventTarget();

  assert.equal(isDialogBackdropClick({ currentTarget: dialog, target: content }), false);
  assert.equal(isDialogBackdropClick({ currentTarget: dialog, target: null }), false);
});
