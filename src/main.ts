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
import { formatWorkbookWindowTitle } from "./window-title";
import { WorkbookController } from "./workbook-controller";
import type {
  ApplyTransactionRequest,
  CellDataRequest,
  CutRangeRequest,
  SheetRangeRequest,
  WorkbookFileOperationResult,
} from "./workbook-core";

const APP_DISPLAY_NAME = "Spready";
const DEFAULT_EXPORT_FILE_NAME = "Sheet1.csv";
const DEFAULT_WORKBOOK_FILE_NAME = "Workbook.spready";
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
  canCut: boolean;
  canDelete: boolean;
};

type SaveWorkbookFileAsArgs = {
  defaultPath?: string;
};

type UnsavedChangesResolution = "cancel" | "discard" | "none" | "save";

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
  const title = formatWorkbookWindowTitle(summary, APP_DISPLAY_NAME);

  for (const browserWindow of BrowserWindow.getAllWindows()) {
    browserWindow.setTitle(title);
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
      enabled: args.canCut,
      label: "Cut",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.cut, browserWindow);
      },
    },
    {
      enabled: args.canCut,
      label: "Cut Values",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.cutValues, browserWindow);
      },
    },
    { type: "separator" },
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

async function chooseWorkbookSavePath(
  browserWindow?: BrowserWindow | null,
  defaultPath?: string,
) {
  const targetWindow = getTargetWindow(browserWindow);
  const dialogOptions: SaveDialogOptions = {
    title: "Save Workbook",
    defaultPath: defaultPath ?? DEFAULT_WORKBOOK_FILE_NAME,
    filters: [{ name: "Spready Workbooks", extensions: ["spready"] }],
  };
  const result = targetWindow
    ? await dialog.showSaveDialog(targetWindow, dialogOptions)
    : await dialog.showSaveDialog(dialogOptions);

  if (result.canceled || !result.filePath) {
    return null;
  }

  return result.filePath;
}

async function saveCurrentWorkbook(
  browserWindow?: BrowserWindow | null,
  requestedFilePath?: string,
  defaultPath?: string,
): Promise<WorkbookFileOperationResult | null> {
  try {
    const summary = workbookController.getSummary();
    const filePath =
      requestedFilePath ??
      summary.documentFilePath ??
      (await chooseWorkbookSavePath(browserWindow, defaultPath));

    if (!filePath) {
      return null;
    }

    return await workbookController.saveWorkbookFile({ filePath });
  } catch (error) {
    dialog.showErrorBox(
      "Save workbook failed",
      error instanceof Error
        ? error.message
        : "The workbook file could not be saved.",
    );

    return null;
  }
}

async function resolveUnsavedChanges(
  browserWindow?: BrowserWindow | null,
): Promise<UnsavedChangesResolution> {
  const summary = workbookController.getSummary();

  if (!summary.hasUnsavedChanges) {
    return "none";
  }

  const targetWindow = getTargetWindow(browserWindow);
  const options = {
    type: "warning" as const,
    buttons: ["Save", "Discard", "Cancel"],
    cancelId: 2,
    defaultId: 0,
    noLink: true,
    title: "Unsaved Changes",
    message: "Save the current workbook before continuing?",
    detail: "Your unsaved changes will be lost if you discard them.",
  };
  const result = targetWindow
    ? await dialog.showMessageBox(targetWindow, options)
    : await dialog.showMessageBox(options);

  if (result.response === 0) {
    return (await saveCurrentWorkbook(browserWindow, undefined, undefined))
      ? "save"
      : "cancel";
  }

  if (result.response === 1) {
    return "discard";
  }

  return "cancel";
}

async function createNewWorkbookWithPrompt(
  browserWindow?: BrowserWindow | null,
) {
  try {
    const unsavedChangesResolution = await resolveUnsavedChanges(browserWindow);

    if (unsavedChangesResolution === "cancel") {
      return;
    }

    workbookController.createNewWorkbook({
      discardUnsavedChanges: unsavedChangesResolution === "discard",
    });
  } catch (error) {
    dialog.showErrorBox(
      "New workbook failed",
      error instanceof Error
        ? error.message
        : "The new workbook could not be created.",
    );
  }
}

function buildAppMenu() {
  const template: MenuItemConstructorOptions[] = [
    {
      label: "File",
      submenu: [
        {
          label: "New Workbook",
          accelerator: "CmdOrCtrl+N",
          click: () => {
            void createNewWorkbookWithPrompt();
          },
        },
        { type: "separator" },
        {
          label: "Open Workbook",
          accelerator: "CmdOrCtrl+O",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.openWorkbook);
          },
        },
        {
          label: "Save Workbook",
          accelerator: "CmdOrCtrl+S",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.saveWorkbook);
          },
        },
        {
          label: "Save Workbook As",
          accelerator: "CmdOrCtrl+Shift+S",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.saveWorkbookAs);
          },
        },
        { type: "separator" },
        {
          label: "Import CSV",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.importCsv);
          },
        },
        {
          label: "Export CSV",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.exportCsv);
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
          accelerator: "CmdOrCtrl+X",
          label: "Cut",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.cut);
          },
        },
        {
          accelerator: "CmdOrCtrl+Shift+X",
          label: "Cut Values",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.cutValues);
          },
        },
        { type: "separator" },
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
  let isClosePromptPending = false;
  let isCloseAuthorized = false;
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
  mainWindow.setTitle(
    formatWorkbookWindowTitle(
      workbookController.getSummary(),
      APP_DISPLAY_NAME,
    ),
  );

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

  mainWindow.on("close", (event) => {
    if (isCloseAuthorized) {
      return;
    }

    event.preventDefault();

    if (isClosePromptPending) {
      return;
    }

    isClosePromptPending = true;

    void resolveUnsavedChanges(mainWindow)
      .then((resolution) => {
        if (resolution === "cancel") {
          return;
        }

        isCloseAuthorized = true;
        mainWindow.close();
      })
      .finally(() => {
        isClosePromptPending = false;
      });
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

ipcMain.handle("dialog:open-workbook-file", async (event) => {
  try {
    const browserWindow = BrowserWindow.fromWebContents(event.sender);
    const unsavedChangesResolution = await resolveUnsavedChanges(browserWindow);

    if (unsavedChangesResolution === "cancel") {
      return { canceled: true as const };
    }

    const targetWindow = getTargetWindow(browserWindow);
    const dialogOptions: OpenDialogOptions = {
      title: "Open Workbook",
      properties: ["openFile"],
      filters: [{ name: "Spready Workbooks", extensions: ["spready"] }],
    };
    const result = targetWindow
      ? await dialog.showOpenDialog(targetWindow, dialogOptions)
      : await dialog.showOpenDialog(dialogOptions);

    if (result.canceled || result.filePaths.length === 0) {
      return { canceled: true as const };
    }

    return {
      canceled: false as const,
      ...(await workbookController.openWorkbookFile({
        discardUnsavedChanges: unsavedChangesResolution === "discard",
        filePath: result.filePaths[0],
      })),
    };
  } catch (error) {
    dialog.showErrorBox(
      "Open workbook failed",
      error instanceof Error
        ? error.message
        : "The workbook file could not be opened.",
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
  "dialog:save-workbook-file-as",
  async (event, args?: SaveWorkbookFileAsArgs) => {
    const browserWindow = BrowserWindow.fromWebContents(event.sender);
    const filePath = await chooseWorkbookSavePath(
      browserWindow,
      args?.defaultPath,
    );

    if (!filePath) {
      return { canceled: true as const };
    }

    const result = await saveCurrentWorkbook(browserWindow, filePath);

    if (!result) {
      return { canceled: true as const };
    }

    return {
      canceled: false as const,
      ...result,
    };
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

ipcMain.handle("workbook:cut-range", (_event, args: CutRangeRequest) =>
  workbookController.cutRange(args),
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

ipcMain.handle("workbook:save-file", (_event, args: { filePath: string }) =>
  workbookController.saveWorkbookFile(args),
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
