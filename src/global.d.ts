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
  RecentWorkbooksResult,
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

declare global {
  interface Window {
    appShell: {
      applyTransaction: (request: ApplyTransactionRequest) => Promise<ApplyTransactionResult>;
      copyRange: (request: CopyRangeRequest) => Promise<CopyRangeResult>;
      cutRange: (request: CutRangeRequest) => Promise<CutRangeResult>;
      getUndoTree: () => Promise<WorkbookUndoTree>;
      getCellData: (request: CellDataRequest) => Promise<CellDataResult>;
      getChart: (chartId: string) => Promise<WorkbookChartResult>;
      getChartPreview: (chartId: string) => Promise<WorkbookChartPreview>;
      getSheetCsv: (sheetId?: string) => Promise<string>;
      getSheetChartPreviews: (sheetId?: string) => Promise<WorkbookSheetChartPreviewsResult>;
      getSheetCharts: (sheetId?: string) => Promise<WorkbookSheetChartsResult>;
      getSheetDisplayRange: (request: SheetRangeRequest) => Promise<SheetDisplayRangeResult>;
      getSheetRange: (request: SheetRangeRequest) => Promise<SheetRangeResult>;
      getSheetStyleRange: (request: SheetRangeRequest) => Promise<SheetStyleRangeResult>;
      getUsedRange: (sheetId?: string) => Promise<UsedRangeResult>;
      getWorkbookSummary: () => Promise<WorkbookSummary>;
      getInstallerStatus: () => Promise<InstallerStatus>;
      installCurrentApp: (options: InstallerOptions) => Promise<InstallerOperationResult>;
      applyInstallerOptions: (options: InstallerOptions) => Promise<InstallerOperationResult>;
      startUninstall: () => Promise<InstallerOperationResult>;
      checkForInstallerUpdates: (
        request?: InstallerCheckUpdatesRequest,
      ) => Promise<InstallerCheckUpdatesResult>;
      logStartupTiming: (event: string, detail?: string) => void;
      name: string;
      readClipboard: () => Promise<ClipboardReadResult>;
      onMenuAction: (listener: (action: AppMenuAction) => void) => () => void;
      onWorkbookChanged: (listener: (summary: WorkbookSummary) => void) => () => void;
      openCsvFile: () => Promise<OpenCsvFileResult>;
      openWorkbookFile: () => Promise<OpenWorkbookFileResult>;
      getRecentWorkbooks: () => Promise<RecentWorkbooksResult>;
      addRecentWorkbook: (filePath: string) => Promise<RecentWorkbooksResult>;
      removeRecentWorkbook: (filePath: string) => Promise<RecentWorkbooksResult>;
      clearRecentWorkbooks: () => Promise<RecentWorkbooksResult>;
      pasteRange: (request: PasteRangeRequest) => Promise<ApplyTransactionResult>;
      saveCsvFile: (content: string, defaultPath?: string) => Promise<SaveCsvFileResult>;
      setChartDialogOpen: (isOpen: boolean) => Promise<void>;
      showCellContextMenu: (request: {
        canCopy: boolean;
        canCut: boolean;
        canDelete: boolean;
        canFormat: boolean;
        canDeleteTable?: boolean;
        canInsertTable?: boolean;
        canSortTable?: boolean;
      }) => Promise<void>;
      redo: (request?: WorkbookRedoRequest) => Promise<WorkbookHistoryResult>;
      undo: (request?: WorkbookHistoryRequest) => Promise<WorkbookHistoryResult>;
      writeClipboard: (request: ClipboardWriteRequest) => Promise<void>;
      saveWorkbookFile: (filePath: string) => Promise<WorkbookFileOperationResult>;
      saveWorkbookFileAs: (defaultPath?: string) => Promise<SaveWorkbookFileAsResult>;
    };
  }
}

export {};
