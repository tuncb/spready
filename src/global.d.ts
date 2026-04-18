type AppMenuAction = 'import' | 'export';

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
      name: string;
      onMenuAction: (listener: (action: AppMenuAction) => void) => () => void;
      openCsvFile: () => Promise<OpenCsvFileResult>;
      saveCsvFile: (content: string, defaultPath?: string) => Promise<SaveCsvFileResult>;
    };
  }
}

export {};
