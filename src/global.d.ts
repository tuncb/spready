import type { AppMenuAction } from "./app-menu";
import type { ClipboardReadResult, ClipboardWriteRequest } from "./clipboard";
import type {
  ApplyTransactionRequest,
  ApplyTransactionResult,
  CellDataRequest,
  CellDataResult,
  SheetDisplayRangeResult,
  SheetRangeRequest,
  SheetRangeResult,
  UsedRangeResult,
  WorkbookFileOperationResult,
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
      applyTransaction: (
        request: ApplyTransactionRequest,
      ) => Promise<ApplyTransactionResult>;
      getCellData: (request: CellDataRequest) => Promise<CellDataResult>;
      getSheetCsv: (sheetId?: string) => Promise<string>;
      getSheetDisplayRange: (
        request: SheetRangeRequest,
      ) => Promise<SheetDisplayRangeResult>;
      getSheetRange: (request: SheetRangeRequest) => Promise<SheetRangeResult>;
      getUsedRange: (sheetId?: string) => Promise<UsedRangeResult>;
      getWorkbookSummary: () => Promise<WorkbookSummary>;
      name: string;
      readClipboard: () => Promise<ClipboardReadResult>;
      onMenuAction: (listener: (action: AppMenuAction) => void) => () => void;
      onWorkbookChanged: (
        listener: (summary: WorkbookSummary) => void,
      ) => () => void;
      openCsvFile: () => Promise<OpenCsvFileResult>;
      openWorkbookFile: () => Promise<OpenWorkbookFileResult>;
      saveCsvFile: (
        content: string,
        defaultPath?: string,
      ) => Promise<SaveCsvFileResult>;
      showCellContextMenu: (request: {
        canCopy: boolean;
        canDelete: boolean;
      }) => Promise<void>;
      writeClipboard: (request: ClipboardWriteRequest) => Promise<void>;
      saveWorkbookFile: (
        filePath: string,
      ) => Promise<WorkbookFileOperationResult>;
      saveWorkbookFileAs: (
        defaultPath?: string,
      ) => Promise<SaveWorkbookFileAsResult>;
    };
  }
}

export {};
