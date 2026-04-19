export const SPREADY_CLIPBOARD_FORMAT = "application/x-spready-cells+json";

export interface SpreadyClipboardPayload {
  displayText: string;
  displayValues: string[][];
  rawText: string;
  rawValues: string[][];
}

export interface ClipboardReadResult {
  payload?: SpreadyClipboardPayload;
  text: string;
}

export interface ClipboardWriteRequest {
  payload?: SpreadyClipboardPayload;
  text: string;
}
