import type { ClipboardRangePayload } from "./workbook-core";

export const SPREADY_CLIPBOARD_FORMAT = "application/x-spready-cells+json";

export type SpreadyClipboardPayload = ClipboardRangePayload;

export interface ClipboardReadResult {
  payload?: SpreadyClipboardPayload;
  text: string;
}

export interface ClipboardWriteRequest {
  payload?: SpreadyClipboardPayload;
  text: string;
}
