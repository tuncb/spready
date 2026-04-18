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

declare global {
  interface Window {
    appShell: {
      applyTransaction: (
        request: ApplyTransactionRequest,
      ) => Promise<ApplyTransactionResult>;
      getCellData: (request: CellDataRequest) => Promise<CellDataResult>;
      getControlInfo: () => Promise<ControlServerInfo>;
      getSheetCsv: (sheetId?: string) => Promise<string>;
      getSheetDisplayRange: (
        request: SheetRangeRequest,
      ) => Promise<SheetDisplayRangeResult>;
      getSheetRange: (request: SheetRangeRequest) => Promise<SheetRangeResult>;
      getUsedRange: (sheetId?: string) => Promise<UsedRangeResult>;
      getWorkbookSummary: () => Promise<WorkbookSummary>;
      name: string;
      onMenuAction: (listener: (action: AppMenuAction) => void) => () => void;
      onWorkbookChanged: (
        listener: (summary: WorkbookSummary) => void,
      ) => () => void;
      openCsvFile: () => Promise<OpenCsvFileResult>;
      saveCsvFile: (
        content: string,
        defaultPath?: string,
      ) => Promise<SaveCsvFileResult>;
    };
  }
}

export {};
