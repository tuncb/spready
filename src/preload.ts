import { contextBridge, ipcRenderer } from "electron";

import type { AppMenuAction } from "./app-menu";
import type {
  ApplyTransactionRequest,
  ApplyTransactionResult,
  CellDataRequest,
  CellDataResult,
  ControlServerInfo,
  SheetDisplayRangeResult,
  SheetRangeRequest,
  SheetRangeResult,
  UsedRangeResult,
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
  getControlInfo: () =>
    ipcRenderer.invoke("control:get-info") as Promise<ControlServerInfo>,
  getSheetCsv: (sheetId?: string) =>
    ipcRenderer.invoke("workbook:get-sheet-csv", {
      sheetId,
    }) as Promise<string>,
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
  saveCsvFile: (content: string, defaultPath?: string) =>
    ipcRenderer.invoke("dialog:save-csv-file", {
      content,
      defaultPath,
    }) as Promise<SaveCsvFileResult>,
});
