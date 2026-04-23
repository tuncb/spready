import { contextBridge, ipcRenderer } from "electron";

import type { AppMenuAction } from "./app-menu";
import type { ClipboardReadResult, ClipboardWriteRequest } from "./clipboard";
import type {
  ApplyTransactionRequest,
  ApplyTransactionResult,
  CellDataRequest,
  CellDataResult,
  CutRangeRequest,
  CutRangeResult,
  WorkbookChartPreview,
  WorkbookChartResult,
  SheetDisplayRangeResult,
  SheetRangeRequest,
  SheetRangeResult,
  UsedRangeResult,
  WorkbookFileOperationResult,
  WorkbookSheetChartsResult,
  WorkbookSummary,
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

contextBridge.exposeInMainWorld("appShell", {
  applyTransaction: (request: ApplyTransactionRequest) =>
    ipcRenderer.invoke(
      "workbook:apply-transaction",
      request,
    ) as Promise<ApplyTransactionResult>,
  getCellData: (request: CellDataRequest) =>
    ipcRenderer.invoke(
      "workbook:get-cell-data",
      request,
    ) as Promise<CellDataResult>,
  getChart: (chartId: string) =>
    ipcRenderer.invoke("workbook:get-chart", {
      chartId,
    }) as Promise<WorkbookChartResult>,
  getChartPreview: (chartId: string) =>
    ipcRenderer.invoke("workbook:get-chart-preview", {
      chartId,
    }) as Promise<WorkbookChartPreview>,
  cutRange: (request: CutRangeRequest) =>
    ipcRenderer.invoke(
      "workbook:cut-range",
      request,
    ) as Promise<CutRangeResult>,
  getSheetCsv: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-sheet-csv", {
      sheetId,
    }) as Promise<string>,
  getSheetCharts: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-sheet-charts", {
      sheetId,
    }) as Promise<WorkbookSheetChartsResult>,
  getSheetDisplayRange: (request: SheetRangeRequest) =>
    ipcRenderer.invoke(
      "workbook:get-display-range",
      request,
    ) as Promise<SheetDisplayRangeResult>,
  getSheetRange: (request: SheetRangeRequest) =>
    ipcRenderer.invoke(
      "workbook:get-range",
      request,
    ) as Promise<SheetRangeResult>,
  getUsedRange: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-used-range", {
      sheetId,
    }) as Promise<UsedRangeResult>,
  getWorkbookSummary: () =>
    ipcRenderer.invoke("workbook:get-summary") as Promise<WorkbookSummary>,
  name: "Spready",
  readClipboard: () =>
    ipcRenderer.invoke("clipboard:read") as Promise<ClipboardReadResult>,
  onMenuAction: (listener: (action: AppMenuAction) => void) => {
    const wrappedListener = (
      _event: Electron.IpcRendererEvent,
      action: AppMenuAction,
    ) => {
      listener(action);
    };

    ipcRenderer.on("app-menu:action", wrappedListener);

    return () => {
      ipcRenderer.off("app-menu:action", wrappedListener);
    };
  },
  onWorkbookChanged: (listener: (summary: WorkbookSummary) => void) => {
    const wrappedListener = (
      _event: Electron.IpcRendererEvent,
      summary: WorkbookSummary,
    ) => {
      listener(summary);
    };

    ipcRenderer.on("workbook:changed", wrappedListener);

    return () => {
      ipcRenderer.off("workbook:changed", wrappedListener);
    };
  },
  openCsvFile: () =>
    ipcRenderer.invoke("dialog:open-csv-file") as Promise<OpenCsvFileResult>,
  openWorkbookFile: () =>
    ipcRenderer.invoke(
      "dialog:open-workbook-file",
    ) as Promise<OpenWorkbookFileResult>,
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
