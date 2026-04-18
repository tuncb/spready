import { contextBridge, ipcRenderer } from 'electron';

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

contextBridge.exposeInMainWorld('appShell', {
  name: 'Spready',
  onMenuAction: (listener: (action: AppMenuAction) => void) => {
    const wrappedListener = (_event: Electron.IpcRendererEvent, action: AppMenuAction) => {
      listener(action);
    };

    ipcRenderer.on('app-menu:action', wrappedListener);

    return () => {
      ipcRenderer.off('app-menu:action', wrappedListener);
    };
  },
  openCsvFile: () => ipcRenderer.invoke('dialog:open-csv-file') as Promise<OpenCsvFileResult>,
  saveCsvFile: (content: string, defaultPath?: string) =>
    ipcRenderer.invoke('dialog:save-csv-file', { content, defaultPath }) as Promise<SaveCsvFileResult>,
});
