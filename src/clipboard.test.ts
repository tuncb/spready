import assert from "node:assert/strict";
import test from "node:test";

import { createSpreadyClipboardHtml, parseSpreadyClipboardHtmlPayload } from "./clipboard";
import type { ClipboardRangePayload } from "./workbook-core";

test("Spready clipboard HTML preserves visible text and structured payload", () => {
  const payload: ClipboardRangePayload = {
    displayText: "1\t2",
    displayValues: [["1", "2"]],
    rawText: "1\t=A1+1",
    rawValues: [["1", "=A1+1"]],
    styles: [
      {
        columnOffset: 1,
        rowOffset: 0,
        style: {
          bold: true,
        },
      },
    ],
    tables: [],
  };

  const html = createSpreadyClipboardHtml('1\t<2 & "3"', payload);

  assert.match(html, /<pre>1\t&lt;2 &amp; &quot;3&quot;<\/pre>/u);
  assert.deepEqual(parseSpreadyClipboardHtmlPayload(html), payload);
});

test("Spready clipboard HTML parser ignores missing or invalid payloads", () => {
  assert.equal(parseSpreadyClipboardHtmlPayload("<pre>plain text</pre>"), undefined);
  assert.equal(parseSpreadyClipboardHtmlPayload("<!--spready-cells:not-base64-->"), undefined);
});
