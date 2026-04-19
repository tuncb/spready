import path from "node:path";

import type { WorkbookSummary } from "./workbook-core";

const UNTITLED_WORKBOOK_LABEL = "Untitled";

export function formatWorkbookWindowTitle(
  summary: WorkbookSummary,
  appName: string,
): string {
  const fileName = summary.documentFilePath
    ? path.basename(summary.documentFilePath)
    : UNTITLED_WORKBOOK_LABEL;
  const dirtyPrefix = summary.hasUnsavedChanges ? "*" : "";

  return `${appName} - ${dirtyPrefix}${fileName}`;
}
