import type { ClipboardRangePayload } from "./workbook-core";

export const SPREADY_CLIPBOARD_FORMAT = "application/x-spready-cells+json";
const SPREADY_CLIPBOARD_HTML_PAYLOAD_PREFIX = "spready-cells:";
const SPREADY_CLIPBOARD_HTML_PAYLOAD_PATTERN = /<!--\s*spready-cells:([A-Za-z0-9+/=]+)\s*-->/u;

export type SpreadyClipboardPayload = ClipboardRangePayload;

export interface ClipboardReadResult {
  payload?: SpreadyClipboardPayload;
  text: string;
}

export interface ClipboardWriteRequest {
  payload?: SpreadyClipboardPayload;
  text: string;
}

function escapeHtmlText(text: string): string {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

export function createSpreadyClipboardHtml(text: string, payload: SpreadyClipboardPayload): string {
  const encodedPayload = Buffer.from(JSON.stringify(payload), "utf8").toString("base64");

  return [
    "<!doctype html>",
    '<html><head><meta charset="utf-8"></head><body>',
    `<pre>${escapeHtmlText(text)}</pre>`,
    `<!--${SPREADY_CLIPBOARD_HTML_PAYLOAD_PREFIX}${encodedPayload}-->`,
    "</body></html>",
  ].join("");
}

export function parseSpreadyClipboardHtmlPayload(
  html: string,
): SpreadyClipboardPayload | undefined {
  const match = SPREADY_CLIPBOARD_HTML_PAYLOAD_PATTERN.exec(html);

  if (!match) {
    return undefined;
  }

  try {
    return JSON.parse(Buffer.from(match[1], "base64").toString("utf8")) as SpreadyClipboardPayload;
  } catch {
    return undefined;
  }
}
