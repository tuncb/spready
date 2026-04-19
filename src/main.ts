import { promises as fs } from "node:fs";
import path from "node:path";

import {
  app,
  BrowserWindow,
  clipboard,
  dialog,
  ipcMain,
  Menu,
  type MenuItemConstructorOptions,
  type OpenDialogOptions,
  type SaveDialogOptions,
} from "electron";
import started from "electron-squirrel-startup";

import { APP_MENU_ACTIONS, type AppMenuAction } from "./app-menu";
import {
  SPREADY_CLIPBOARD_FORMAT,
  type ClipboardReadResult,
  type ClipboardWriteRequest,
  type SpreadyClipboardPayload,
} from "./clipboard";
import { SpreadyControlServer } from "./control-server";
import {
  clearDiscoveredControlInfo,
  writeDiscoveredControlInfo,
} from "./control-discovery";
import { WorkbookController } from "./workbook-controller";
import type {
  ApplyTransactionRequest,
  CellDataRequest,
  SheetRangeRequest,
} from "./workbook-core";

const APP_DISPLAY_NAME = "Spready";
const DEFAULT_EXPORT_FILE_NAME = "Sheet1.csv";
const DEFAULT_CONTROL_HOST = "127.0.0.1";
const DEFAULT_CONTROL_PORT = 45731;

const workbookController = new WorkbookController();
const configuredControlPort = Number.parseInt(
  process.env.SPREADY_CONTROL_PORT ?? `${DEFAULT_CONTROL_PORT}`,
  10,
);
const controlServer = new SpreadyControlServer(
  workbookController,
  DEFAULT_CONTROL_HOST,
  Number.isNaN(configuredControlPort)
    ? DEFAULT_CONTROL_PORT
    : configuredControlPort,
);

if (started) {
  app.quit();
}

app.setName(APP_DISPLAY_NAME);

type SaveCsvFileArgs = {
  content: string;
  defaultPath?: string;
};

type ShowCellContextMenuArgs = {
  canCopy: boolean;
  canDelete: boolean;
};

function readSpreadyClipboardPayload(): SpreadyClipboardPayload | undefined {
  const buffer = clipboard.readBuffer(SPREADY_CLIPBOARD_FORMAT);

  if (buffer.length === 0) {
    return undefined;
  }

  try {
    return JSON.parse(buffer.toString("utf8")) as SpreadyClipboardPayload;
  } catch {
    return undefined;
  }
}

function getTargetWindow(
  browserWindow?: BrowserWindow | null,
): BrowserWindow | null {
  return (
    browserWindow ??
    BrowserWindow.getAllWindows().find((window) => window.isFocused()) ??
    BrowserWindow.getAllWindows()[0] ??
    null
  );
}

function sendMenuAction(
  action: AppMenuAction,
  browserWindow?: BrowserWindow | null,
) {
  getTargetWindow(browserWindow)?.webContents.send("app-menu:action", action);
}

function broadcastWorkbookChanged() {
  const summary = workbookController.getSummary();

  for (const browserWindow of BrowserWindow.getAllWindows()) {
    browserWindow.webContents.send("workbook:changed", summary);
  }
}

async function showAboutDialog(browserWindow?: BrowserWindow | null) {
  const targetWindow = getTargetWindow(browserWindow);
  const controlInfo = controlServer.getInfo();
  const options = {
    type: "info" as const,
    buttons: ["OK"],
    title: `About ${APP_DISPLAY_NAME}`,
    message: APP_DISPLAY_NAME,
    detail: `Version ${app.getVersion()}\n\ntcp://${controlInfo.host}:${controlInfo.port}`,
  };

  if (targetWindow) {
    await dialog.showMessageBox(targetWindow, options);
    return;
  }

  await dialog.showMessageBox(options);
}

function buildCellContextMenu(
  browserWindow: BrowserWindow,
  args: ShowCellContextMenuArgs,
) {
  return Menu.buildFromTemplate([
    {
      enabled: args.canCopy,
      label: "Copy",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.copy, browserWindow);
      },
    },
    {
      enabled: args.canCopy,
      label: "Copy Values",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.copyValues, browserWindow);
      },
    },
    { type: "separator" },
    {
      label: "Paste",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.paste, browserWindow);
      },
    },
    {
      label: "Paste Values",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.pasteValues, browserWindow);
      },
    },
    { type: "separator" },
    {
      enabled: args.canDelete,
      label: "Delete",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.deleteSelection, browserWindow);
      },
    },
  ]);
}

function buildAppMenu() {
  const template: MenuItemConstructorOptions[] = [
    {
      label: "File",
      submenu: [
        {
          label: "Import",
          accelerator: "CmdOrCtrl+O",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.import);
          },
        },
        {
          label: "Export",
          accelerator: "CmdOrCtrl+Shift+S",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.export);
          },
        },
        { type: "separator" },
        {
          label: "Exit",
          accelerator: "Alt+F4",
          click: () => {
            app.quit();
          },
        },
      ],
    },
    {
      label: "Edit",
      submenu: [
        {
          accelerator: "CmdOrCtrl+C",
          label: "Copy",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.copy);
          },
        },
        {
          accelerator: "CmdOrCtrl+Shift+C",
          label: "Copy Values",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.copyValues);
          },
        },
        { type: "separator" },
        {
          accelerator: "CmdOrCtrl+V",
          label: "Paste",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.paste);
          },
        },
        {
          accelerator: "CmdOrCtrl+Shift+V",
          label: "Paste Values",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.pasteValues);
          },
        },
        { type: "separator" },
        {
          label: "Delete",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.deleteSelection);
          },
        },
      ],
    },
    {
      label: "Sheet",
      submenu: [
        {
          label: "Add Row",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.addRow);
          },
        },
        {
          label: "Add Column",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.addColumn);
          },
        },
        { type: "separator" },
        {
          label: "New Sheet",
          accelerator: "CmdOrCtrl+Shift+N",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.newSheet);
          },
        },
        {
          label: "Delete Sheet",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.deleteSheet);
          },
        },
      ],
    },
    {
      label: "Help",
      submenu: [
        {
          label: "About",
          click: () => {
            void showAboutDialog();
          },
        },
      ],
    },
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

const createWindow = () => {
  const mainWindow = new BrowserWindow({
    width: 960,
    height: 640,
    minWidth: 720,
    minHeight: 480,
    show: false,
    autoHideMenuBar: false,
    backgroundColor: "#f3efe8",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      sandbox: true,
    },
  });

  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL);
  } else {
    mainWindow.loadFile(
      path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`),
    );
  }

  mainWindow.once("ready-to-show", () => {
    mainWindow.show();
  });

  return mainWindow;
};

ipcMain.handle("dialog:open-csv-file", async (event) => {
  try {
    const browserWindow = BrowserWindow.fromWebContents(event.sender);
    const targetWindow = getTargetWindow(browserWindow);
    const dialogOptions: OpenDialogOptions = {
      title: "Import CSV",
      properties: ["openFile"],
      filters: [{ name: "CSV Files", extensions: ["csv"] }],
    };
    const result = targetWindow
      ? await dialog.showOpenDialog(targetWindow, dialogOptions)
      : await dialog.showOpenDialog(dialogOptions);

    if (result.canceled || result.filePaths.length === 0) {
      return { canceled: true as const };
    }

    const filePath = result.filePaths[0];
    const content = await fs.readFile(filePath, "utf8");

    return {
      canceled: false as const,
      content,
      filePath,
    };
  } catch (error) {
    dialog.showErrorBox(
      "Import failed",
      error instanceof Error
        ? error.message
        : "The CSV file could not be opened.",
    );

    return { canceled: true as const };
  }
});

ipcMain.handle("dialog:save-csv-file", async (event, args: SaveCsvFileArgs) => {
  try {
    const browserWindow = BrowserWindow.fromWebContents(event.sender);
    const targetWindow = getTargetWindow(browserWindow);
    const dialogOptions: SaveDialogOptions = {
      title: "Export CSV",
      defaultPath: args.defaultPath ?? DEFAULT_EXPORT_FILE_NAME,
      filters: [{ name: "CSV Files", extensions: ["csv"] }],
    };
    const saveDialogResult = targetWindow
      ? await dialog.showSaveDialog(targetWindow, dialogOptions)
      : await dialog.showSaveDialog(dialogOptions);

    if (saveDialogResult.canceled || !saveDialogResult.filePath) {
      return { canceled: true as const };
    }

    const filePath = saveDialogResult.filePath.toLowerCase().endsWith(".csv")
      ? saveDialogResult.filePath
      : `${saveDialogResult.filePath}.csv`;

    await fs.writeFile(filePath, args.content, "utf8");

    return {
      canceled: false as const,
      filePath,
    };
  } catch (error) {
    dialog.showErrorBox(
      "Export failed",
      error instanceof Error
        ? error.message
        : "The CSV file could not be saved.",
    );

    return { canceled: true as const };
  }
});

ipcMain.handle("clipboard:read", () => {
  const result: ClipboardReadResult = {
    payload: readSpreadyClipboardPayload(),
    text: clipboard.readText(),
  };

  return result;
});

ipcMain.handle("clipboard:write", (_event, request: ClipboardWriteRequest) => {
  clipboard.clear();
  clipboard.writeText(request.text);

  if (!request.payload) {
    return;
  }

  clipboard.writeBuffer(
    SPREADY_CLIPBOARD_FORMAT,
    Buffer.from(JSON.stringify(request.payload), "utf8"),
  );
});

ipcMain.handle(
  "menu:show-cell-context-menu",
  async (event, args: ShowCellContextMenuArgs) => {
    const browserWindow = BrowserWindow.fromWebContents(event.sender);

    if (!browserWindow) {
      return;
    }

    buildCellContextMenu(browserWindow, args).popup({
      window: browserWindow,
    });
  },
);

ipcMain.handle(
  "workbook:apply-transaction",
  (_event, args: ApplyTransactionRequest) =>
    workbookController.applyTransaction(args),
);

ipcMain.handle("workbook:get-cell-data", (_event, args: CellDataRequest) =>
  workbookController.getCellData(args),
);

ipcMain.handle(
  "workbook:get-display-range",
  (_event, args: SheetRangeRequest) =>
    workbookController.getSheetDisplayRange(args),
);

ipcMain.handle("workbook:get-range", (_event, args: SheetRangeRequest) =>
  workbookController.getSheetRange(args),
);

ipcMain.handle(
  "workbook:get-sheet-csv",
  (_event, args?: { sheetId?: string }) =>
    workbookController.getSheetCsv(args?.sheetId),
);

ipcMain.handle("workbook:get-summary", () => workbookController.getSummary());

ipcMain.handle(
  "workbook:get-used-range",
  (_event, args?: { sheetId?: string }) =>
    workbookController.getUsedRange(args?.sheetId),
);

workbookController.on("changed", () => {
  broadcastWorkbookChanged();
});

app.whenReady().then(() => {
  void controlServer
    .start()
    .then(() => {
      const controlInfo = controlServer.getInfo();
      void writeDiscoveredControlInfo(APP_DISPLAY_NAME, controlInfo);
      console.log(
        `${APP_DISPLAY_NAME} control server listening on tcp://${controlInfo.host}:${controlInfo.port}`,
      );
    })
    .catch((error) => {
      console.error(
        `${APP_DISPLAY_NAME} control server failed to start: ${
          error instanceof Error ? error.message : "unknown error"
        }`,
      );
    });

  createWindow();
  buildAppMenu();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("before-quit", () => {
  void clearDiscoveredControlInfo();
  void controlServer.stop();
});
