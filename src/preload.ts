import { contextBridge, ipcRenderer } from "electron";

import type { AppMenuAction } from "./app-menu";
import type { ClipboardReadResult, ClipboardWriteRequest } from "./clipboard";
import type {
  ApplyTransactionRequest,
  ApplyTransactionResult,
  CellDataRequest,
  CellDataResult,
  CopyRangeRequest,
  CopyRangeResult,
  CutRangeRequest,
  CutRangeResult,
  InstallerCheckUpdatesRequest,
  InstallerCheckUpdatesResult,
  InstallerOperationResult,
  InstallerOptions,
  InstallerStatus,
  PasteRangeRequest,
  SetWorkbookSearchQueryRequest,
  WorkbookChartPreview,
  WorkbookChartResult,
  SheetDisplayRangeResult,
  SheetRangeRequest,
  SheetRangeResult,
  SheetStyleRangeResult,
  UsedRangeResult,
  WorkbookFileOperationResult,
  WorkbookHistoryRequest,
  WorkbookHistoryResult,
  WorkbookRedoRequest,
  WorkbookSheetChartPreviewsResult,
  WorkbookSheetChartsResult,
  WorkbookSearchState,
  WorkbookSummary,
  WorkbookUndoTree,
} from "./workbook-core";

type OpenCsvFileResult =
  | {
      canceled: true;
    }
  | {
      canceled: false;
      content: string;
      filePath: string;
    };

type SaveCsvFileResult =
  | {
      canceled: true;
    }
  | {
      canceled: false;
      filePath: string;
    };

type ShowCellContextMenuRequest = {
  canCopy: boolean;
  canCut: boolean;
  canDelete: boolean;
  canFormat: boolean;
  canDeleteTable?: boolean;
  canInsertTable?: boolean;
  canSortTable?: boolean;
};

type OpenWorkbookFileResult =
  | {
      canceled: true;
    }
  | ({
      canceled: false;
    } & WorkbookFileOperationResult);

type SaveWorkbookFileAsResult =
  | {
      canceled: true;
    }
  | ({
      canceled: false;
    } & WorkbookFileOperationResult);

function getStartupTimingDetail(detail?: string) {
  const rendererMs = `rendererMs=${Math.round(performance.now())}`;

  return detail ? `${rendererMs} ${detail}` : rendererMs;
}

function logStartupTiming(event: string, detail?: string) {
  ipcRenderer.send("startup:timing", {
    detail: getStartupTimingDetail(detail),
    event,
  });
}

logStartupTiming("preload-start");

const MAX_STARTUP_RESOURCE_TIMING_LOGS = 24;
let startupResourceTimingLogCount = 0;

function getResourceTimingDetail(entry: PerformanceResourceTiming) {
  const name = entry.name.length > 120 ? `...${entry.name.slice(-117)}` : entry.name;

  return [
    `name=${name}`,
    `initiatorType=${entry.initiatorType || "unknown"}`,
    `durationMs=${Math.round(entry.duration)}`,
    `transferSize=${entry.transferSize}`,
    `encodedBodySize=${entry.encodedBodySize}`,
  ].join(" ");
}

function logStartupResourceTiming(entry: PerformanceEntry) {
  if (
    startupResourceTimingLogCount >= MAX_STARTUP_RESOURCE_TIMING_LOGS ||
    entry.entryType !== "resource"
  ) {
    return;
  }

  startupResourceTimingLogCount += 1;
  logStartupTiming("resource-timing", getResourceTimingDetail(entry as PerformanceResourceTiming));
}

function logDocumentStartupTiming(event: string) {
  logStartupTiming(event, `readyState=${document.readyState}`);
}

logDocumentStartupTiming("preload-document-state");

document.addEventListener("readystatechange", () => {
  logDocumentStartupTiming("document-readystatechange");
});

document.addEventListener("DOMContentLoaded", () => {
  logDocumentStartupTiming("document-content-loaded");
});

window.addEventListener("load", () => {
  logDocumentStartupTiming("window-load");
});

if (typeof PerformanceObserver !== "undefined") {
  try {
    const resourceObserver = new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) {
        logStartupResourceTiming(entry);
      }
    });

    resourceObserver.observe({
      buffered: true,
      type: "resource",
    });

    window.addEventListener("load", () => {
      resourceObserver.disconnect();
      logStartupTiming("resource-timing-observer-done", `count=${startupResourceTimingLogCount}`);
    });
  } catch {
    logStartupTiming("resource-timing-observer-unavailable");
  }
}

contextBridge.exposeInMainWorld("appShell", {
  applyTransaction: (request: ApplyTransactionRequest) =>
    ipcRenderer.invoke("workbook:apply-transaction", request) as Promise<ApplyTransactionResult>,
  getUndoTree: () => ipcRenderer.invoke("workbook:get-undo-tree") as Promise<WorkbookUndoTree>,
  getSearchState: () =>
    ipcRenderer.invoke("workbook:get-search-state") as Promise<WorkbookSearchState>,
  setSearchQuery: (request: SetWorkbookSearchQueryRequest) =>
    ipcRenderer.invoke("workbook:search-set-query", request) as Promise<WorkbookSearchState>,
  clearSearch: () => ipcRenderer.invoke("workbook:search-clear") as Promise<WorkbookSearchState>,
  goToNextSearchResult: () =>
    ipcRenderer.invoke("workbook:search-next") as Promise<WorkbookSearchState>,
  goToPreviousSearchResult: () =>
    ipcRenderer.invoke("workbook:search-previous") as Promise<WorkbookSearchState>,
  setActiveSearchResult: (index: number) =>
    ipcRenderer.invoke("workbook:search-set-active", {
      index,
    }) as Promise<WorkbookSearchState>,
  redo: (request?: WorkbookRedoRequest) =>
    ipcRenderer.invoke("workbook:redo", request) as Promise<WorkbookHistoryResult>,
  undo: (request?: WorkbookHistoryRequest) =>
    ipcRenderer.invoke("workbook:undo", request) as Promise<WorkbookHistoryResult>,
  getCellData: (request: CellDataRequest) =>
    ipcRenderer.invoke("workbook:get-cell-data", request) as Promise<CellDataResult>,
  getChart: (chartId: string) =>
    ipcRenderer.invoke("workbook:get-chart", {
      chartId,
    }) as Promise<WorkbookChartResult>,
  getChartPreview: (chartId: string) =>
    ipcRenderer.invoke("workbook:get-chart-preview", {
      chartId,
    }) as Promise<WorkbookChartPreview>,
  copyRange: (request: CopyRangeRequest) =>
    ipcRenderer.invoke("workbook:copy-range", request) as Promise<CopyRangeResult>,
  cutRange: (request: CutRangeRequest) =>
    ipcRenderer.invoke("workbook:cut-range", request) as Promise<CutRangeResult>,
  getSheetCsv: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-sheet-csv", {
      sheetId,
    }) as Promise<string>,
  getSheetCharts: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-sheet-charts", {
      sheetId,
    }) as Promise<WorkbookSheetChartsResult>,
  getSheetChartPreviews: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-sheet-chart-previews", {
      sheetId,
    }) as Promise<WorkbookSheetChartPreviewsResult>,
  getSheetDisplayRange: (request: SheetRangeRequest) =>
    ipcRenderer.invoke("workbook:get-display-range", request) as Promise<SheetDisplayRangeResult>,
  getSheetRange: (request: SheetRangeRequest) =>
    ipcRenderer.invoke("workbook:get-range", request) as Promise<SheetRangeResult>,
  getSheetStyleRange: (request: SheetRangeRequest) =>
    ipcRenderer.invoke("workbook:get-style-range", request) as Promise<SheetStyleRangeResult>,
  getUsedRange: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-used-range", {
      sheetId,
    }) as Promise<UsedRangeResult>,
  getWorkbookSummary: () => ipcRenderer.invoke("workbook:get-summary") as Promise<WorkbookSummary>,
  getInstallerStatus: () => ipcRenderer.invoke("installer:get-status") as Promise<InstallerStatus>,
  installCurrentApp: (options: InstallerOptions) =>
    ipcRenderer.invoke(
      "installer:install-current-app",
      options,
    ) as Promise<InstallerOperationResult>,
  applyInstallerOptions: (options: InstallerOptions) =>
    ipcRenderer.invoke("installer:apply-options", options) as Promise<InstallerOperationResult>,
  startUninstall: () =>
    ipcRenderer.invoke("installer:start-uninstall") as Promise<InstallerOperationResult>,
  checkForInstallerUpdates: (request?: InstallerCheckUpdatesRequest) =>
    ipcRenderer.invoke(
      "installer:check-for-updates",
      request,
    ) as Promise<InstallerCheckUpdatesResult>,
  logStartupTiming,
  name: "Spready",
  readClipboard: () => ipcRenderer.invoke("clipboard:read") as Promise<ClipboardReadResult>,
  onMenuAction: (listener: (action: AppMenuAction) => void) => {
    const wrappedListener = (_event: Electron.IpcRendererEvent, action: AppMenuAction) => {
      listener(action);
    };

    ipcRenderer.on("app-menu:action", wrappedListener);

    return () => {
      ipcRenderer.off("app-menu:action", wrappedListener);
    };
  },
  onWorkbookChanged: (listener: (summary: WorkbookSummary) => void) => {
    const wrappedListener = (_event: Electron.IpcRendererEvent, summary: WorkbookSummary) => {
      listener(summary);
    };

    ipcRenderer.on("workbook:changed", wrappedListener);

    return () => {
      ipcRenderer.off("workbook:changed", wrappedListener);
    };
  },
  openCsvFile: () => ipcRenderer.invoke("dialog:open-csv-file") as Promise<OpenCsvFileResult>,
  openWorkbookFile: () =>
    ipcRenderer.invoke("dialog:open-workbook-file") as Promise<OpenWorkbookFileResult>,
  pasteRange: (request: PasteRangeRequest) =>
    ipcRenderer.invoke("workbook:paste-range", request) as Promise<ApplyTransactionResult>,
  saveCsvFile: (content: string, defaultPath?: string) =>
    ipcRenderer.invoke("dialog:save-csv-file", {
      content,
      defaultPath,
    }) as Promise<SaveCsvFileResult>,
  showCellContextMenu: (request: ShowCellContextMenuRequest) =>
    ipcRenderer.invoke("menu:show-cell-context-menu", request) as Promise<void>,
  setChartDialogOpen: (isOpen: boolean) =>
    ipcRenderer.invoke("menu:set-chart-dialog-open", isOpen) as Promise<void>,
  writeClipboard: (request: ClipboardWriteRequest) =>
    ipcRenderer.invoke("clipboard:write", request) as Promise<void>,
  saveWorkbookFile: (filePath: string) =>
    ipcRenderer.invoke("workbook:save-file", {
      filePath,
    }) as Promise<WorkbookFileOperationResult>,
  saveWorkbookFileAs: (defaultPath?: string) =>
    ipcRenderer.invoke("dialog:save-workbook-file-as", {
      defaultPath,
    }) as Promise<SaveWorkbookFileAsResult>,
});

logStartupTiming("preload-bridge-exposed");
