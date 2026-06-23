import { promises as fs } from "node:fs";
import path from "node:path";

import {
  app,
  BrowserWindow,
  clipboard,
  dialog,
  ipcMain,
  Menu,
  shell,
  type MenuItemConstructorOptions,
  type OpenDialogOptions,
  type SaveDialogOptions,
} from "electron";
import started from "electron-squirrel-startup";

import { getMainHelpText, parseMainStartupOptions } from "./app-startup";
import { APP_MENU_ACTIONS, type AppMenuAction } from "./app-menu";
import {
  SPREADY_CLIPBOARD_FORMAT,
  type ClipboardReadResult,
  type ClipboardWriteRequest,
  type SpreadyClipboardPayload,
} from "./clipboard";
import { SpreadyControlServer } from "./control-server";
import { clearDiscoveredControlInfo, writeDiscoveredControlInfo } from "./control-discovery";
import { InstallerService } from "./installer-service";
import { getDefaultRecentWorkbooksFilePath, RecentWorkbooksStore } from "./recent-workbooks";
import { createStartupLogSink, STARTUP_TIMING_LOG_FILE_PATH, StartupTimer } from "./startup-timing";
import { formatWorkbookWindowTitle } from "./window-title";
import { WorkbookController } from "./workbook-controller";
import type {
  ApplyTransactionRequest,
  CellDataRequest,
  ControlAppStatus,
  CopyRangeRequest,
  CutRangeRequest,
  InstallerOptions,
  InstallerCheckUpdatesRequest,
  PasteRangeRequest,
  SheetRangeRequest,
  WorkbookHistoryRequest,
  WorkbookRedoRequest,
  WorkbookFileOperationResult,
} from "./workbook-core";

const APP_DISPLAY_NAME = "Spready";
const APP_ICON_PATH = path.join(__dirname, "..", "..", "assets", "spready.png");
const DEFAULT_EXPORT_FILE_NAME = "Sheet1.csv";
const DEFAULT_WORKBOOK_FILE_NAME = "Workbook.spready";
const DEFAULT_CONTROL_HOST = "127.0.0.1";
const DEFAULT_CONTROL_PORT = 45731;
const MAX_RECENT_WORKBOOK_MENU_LABEL_LENGTH = 72;

const workbookController = new WorkbookController();
const mainStartupOptions = parseMainStartupOptions(process.argv.slice(1));
const configuredControlPort = Number.parseInt(
  process.env.SPREADY_CONTROL_PORT ?? `${DEFAULT_CONTROL_PORT}`,
  10,
);
let isChartDialogOpen = false;
let isAppShowRequested = false;
const isInstallerUninstallCommand = process.argv.includes("--spready-uninstall");
const isConsoleExitMode =
  mainStartupOptions.help || mainStartupOptions.consoleOutputFilePath !== undefined;

if (started) {
  app.quit();
}

app.setName(APP_DISPLAY_NAME);

const installerService = new InstallerService({
  currentAppDirectory: path.dirname(process.execPath),
  currentExecutablePath: process.execPath,
  currentVersion: app.getVersion(),
  isPackaged: app.isPackaged,
  requestQuit: () => {
    setTimeout(() => {
      app.quit();
    }, 100);
  },
  writeShortcut: (shortcutPath, operation, shortcut) =>
    shell.writeShortcutLink(shortcutPath, operation, shortcut),
});

const recentWorkbooksStore = new RecentWorkbooksStore({
  filePath: getDefaultRecentWorkbooksFilePath(process.execPath),
});

const startupTimer = new StartupTimer("spready-main", createStartupLogSink(console.log));
if (!isConsoleExitMode) {
  startupTimer.log(
    "process-start",
    `pid=${process.pid} port=${
      Number.isNaN(configuredControlPort) ? DEFAULT_CONTROL_PORT : configuredControlPort
    } logFile=${STARTUP_TIMING_LOG_FILE_PATH}`,
  );
}

type StartupTimingMessage = {
  detail?: unknown;
  event?: unknown;
};

type SaveCsvFileArgs = {
  content: string;
  defaultPath?: string;
};

type ShowCellContextMenuArgs = {
  canCopy: boolean;
  canCut: boolean;
  canDelete: boolean;
  canFormat: boolean;
  canDeleteTable?: boolean;
  canInsertTable?: boolean;
  canSortTable?: boolean;
};

type SaveWorkbookFileAsArgs = {
  defaultPath?: string;
};

type UnsavedChangesResolution = "cancel" | "discard" | "none" | "save";

function formatStartupTimingMessageDetail(
  browserWindow: BrowserWindow | null,
  detail: string | undefined,
) {
  const windowDetail = browserWindow ? `windowId=${browserWindow.id}` : "windowId=unknown";

  return detail ? `${windowDetail} ${detail}` : windowDetail;
}

function getStartupTimingMessageEvent(message: StartupTimingMessage): string | null {
  if (typeof message.event !== "string") {
    return null;
  }

  if (!/^[a-z0-9-]+$/u.test(message.event)) {
    return null;
  }

  return message.event;
}

function getStartupTimingMessageDetail(message: StartupTimingMessage): string | undefined {
  if (typeof message.detail !== "string") {
    return undefined;
  }

  return message.detail.slice(0, 200);
}

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

function getTargetWindow(browserWindow?: BrowserWindow | null): BrowserWindow | null {
  return (
    browserWindow ??
    BrowserWindow.getAllWindows().find((window) => window.isFocused()) ??
    BrowserWindow.getAllWindows()[0] ??
    null
  );
}

function getControlAppStatus(): ControlAppStatus {
  const windows = BrowserWindow.getAllWindows();
  const visibleWindows = windows.filter((window) => window.isVisible() && !window.isMinimized());
  const readyVisibleWindows = visibleWindows.filter((window) => !window.webContents.isLoading());
  const focusedWindows = windows.filter((window) => window.isFocused());

  return {
    focusedWindowCount: focusedWindows.length,
    frontendVisible: readyVisibleWindows.length > 0,
    visibleWindowCount: visibleWindows.length,
    windowCount: windows.length,
  };
}

function formatControlAppStatus(status: ControlAppStatus) {
  return `frontendVisible=${status.frontendVisible} windowCount=${status.windowCount} visibleWindowCount=${status.visibleWindowCount} focusedWindowCount=${status.focusedWindowCount}`;
}

function flushPendingAppShowRequest(targetWindow: BrowserWindow, reason: string) {
  if (!isAppShowRequested || targetWindow.isDestroyed()) {
    return;
  }

  if (targetWindow.isMinimized()) {
    targetWindow.restore();
  }

  if (targetWindow.webContents.isLoading()) {
    startupTimer.log(
      "show-app-window-deferred",
      `windowId=${targetWindow.id} reason=${reason} loading=true visible=${targetWindow.isVisible()}`,
    );
    return;
  }

  if (!targetWindow.isVisible()) {
    targetWindow.show();
    startupTimer.log("show-app-window-show-called", `windowId=${targetWindow.id} reason=${reason}`);
  }

  if (targetWindow.isVisible()) {
    targetWindow.focus();
    isAppShowRequested = false;
    startupTimer.log(
      "show-app-window-request-satisfied",
      `windowId=${targetWindow.id} reason=${reason}`,
    );
  }
}

function showAppWindow() {
  startupTimer.log("show-app-window-start");
  const targetWindow = BrowserWindow.getAllWindows()[0] ?? createWindow();
  isAppShowRequested = true;

  flushPendingAppShowRequest(targetWindow, "show-app");

  const status = getControlAppStatus();
  startupTimer.log("show-app-window-done", formatControlAppStatus(status));

  return status;
}

const controlServer = new SpreadyControlServer(
  workbookController,
  DEFAULT_CONTROL_HOST,
  Number.isNaN(configuredControlPort) ? DEFAULT_CONTROL_PORT : configuredControlPort,
  {
    getAppStatus: getControlAppStatus,
    showApp: showAppWindow,
    startupTimer,
  },
);

function sendMenuAction(action: AppMenuAction, browserWindow?: BrowserWindow | null) {
  if (isChartDialogOpen) {
    return;
  }

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

function rebuildAppMenuSoon() {
  if (isConsoleExitMode) {
    return;
  }

  buildAppMenu();
}

async function addRecentWorkbook(filePath: string) {
  const result = await recentWorkbooksStore.add(filePath);

  rebuildAppMenuSoon();
  return result;
}

async function removeRecentWorkbook(filePath: string) {
  const result = await recentWorkbooksStore.remove(filePath);

  rebuildAppMenuSoon();
  return result;
}

async function clearRecentWorkbooks() {
  const result = await recentWorkbooksStore.clear();

  rebuildAppMenuSoon();
  return result;
}

function runMenuCommand(command: () => void | Promise<void>) {
  if (isChartDialogOpen) {
    return;
  }

  void command();
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

function buildCellContextMenu(browserWindow: BrowserWindow, args: ShowCellContextMenuArgs) {
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
    { type: "separator" },
    {
      enabled: args.canFormat,
      label: "Format...",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.formatCells, browserWindow);
      },
    },
    {
      enabled: args.canFormat,
      label: "Clear Formatting",
      click: () => {
        sendMenuAction(APP_MENU_ACTIONS.clearFormatting, browserWindow);
      },
    },
    { type: "separator" },
    {
      label: "Table",
      submenu: [
        {
          enabled: (args.canInsertTable ?? false) || (args.canDeleteTable ?? false),
          label: args.canDeleteTable ? "Remove Table" : "Insert Table",
          click: () => {
            sendMenuAction(
              args.canDeleteTable ? APP_MENU_ACTIONS.deleteTable : APP_MENU_ACTIONS.insertTable,
              browserWindow,
            );
          },
        },
        { type: "separator" },
        {
          enabled: args.canSortTable ?? false,
          label: "Sort Ascending",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.sortTableAscending, browserWindow);
          },
        },
        {
          enabled: args.canSortTable ?? false,
          label: "Sort Descending",
          click: () => {
            sendMenuAction(APP_MENU_ACTIONS.sortTableDescending, browserWindow);
          },
        },
      ],
    },
  ]);
}

async function chooseWorkbookSavePath(browserWindow?: BrowserWindow | null, defaultPath?: string) {
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

    const result = await workbookController.saveWorkbookFile({ filePath });

    await addRecentWorkbook(result.filePath);

    return result;
  } catch (error) {
    dialog.showErrorBox(
      "Save workbook failed",
      error instanceof Error ? error.message : "The workbook file could not be saved.",
    );

    return null;
  }
}

async function openWorkbookFilePath(
  filePath: string,
  browserWindow?: BrowserWindow | null,
  unsavedChangesResolution?: UnsavedChangesResolution,
): Promise<WorkbookFileOperationResult | null> {
  try {
    const resolution = unsavedChangesResolution ?? (await resolveUnsavedChanges(browserWindow));

    if (resolution === "cancel") {
      return null;
    }

    const result = await workbookController.openWorkbookFile({
      discardUnsavedChanges: resolution === "discard",
      filePath,
    });

    await addRecentWorkbook(result.filePath);

    return result;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      await removeRecentWorkbook(filePath);
    }

    dialog.showErrorBox(
      "Open workbook failed",
      error instanceof Error ? error.message : "The workbook file could not be opened.",
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
    return (await saveCurrentWorkbook(browserWindow, undefined, undefined)) ? "save" : "cancel";
  }

  if (result.response === 1) {
    return "discard";
  }

  return "cancel";
}

async function createNewWorkbookWithPrompt(browserWindow?: BrowserWindow | null) {
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
      error instanceof Error ? error.message : "The new workbook could not be created.",
    );
  }
}

function truncateMenuLabel(label: string) {
  if (label.length <= MAX_RECENT_WORKBOOK_MENU_LABEL_LENGTH) {
    return label;
  }

  return `...${label.slice(-(MAX_RECENT_WORKBOOK_MENU_LABEL_LENGTH - 3))}`;
}

function formatRecentWorkbookMenuLabel(filePath: string, index: number) {
  const baseName = path.basename(filePath) || filePath;
  const directory = path.dirname(filePath);
  const detail = directory && directory !== "." ? ` (${directory})` : "";

  return `${index + 1}. ${truncateMenuLabel(`${baseName}${detail}`)}`;
}

function buildRecentWorkbooksMenu(): MenuItemConstructorOptions {
  const recentWorkbooks = recentWorkbooksStore.getRecentWorkbooks().workbooks;

  if (recentWorkbooks.length === 0) {
    return {
      enabled: false,
      label: "Open Recent",
      submenu: [
        {
          enabled: false,
          label: "No Recent Workbooks",
        },
      ],
    };
  }

  return {
    label: "Open Recent",
    submenu: [
      ...recentWorkbooks.map(
        (entry, index): MenuItemConstructorOptions => ({
          label: formatRecentWorkbookMenuLabel(entry.filePath, index),
          click: (_menuItem, browserWindow) => {
            runMenuCommand(() => {
              void openWorkbookFilePath(
                entry.filePath,
                browserWindow instanceof BrowserWindow ? browserWindow : undefined,
              );
            });
          },
        }),
      ),
      { type: "separator" },
      {
        label: "Clear Recent Workbooks",
        click: () => {
          runMenuCommand(() => {
            void clearRecentWorkbooks();
          });
        },
      },
    ],
  };
}

function buildAppMenu() {
  const menuEnabled = !isChartDialogOpen;
  const template: MenuItemConstructorOptions[] = [
    {
      enabled: menuEnabled,
      label: "File",
      submenu: [
        {
          label: "New Workbook",
          accelerator: "CmdOrCtrl+N",
          click: () => {
            runMenuCommand(() => createNewWorkbookWithPrompt());
          },
        },
        { type: "separator" },
        {
          label: "Open Workbook...",
          accelerator: "CmdOrCtrl+O",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.openWorkbook);
            });
          },
        },
        buildRecentWorkbooksMenu(),
        {
          label: "Save Workbook",
          accelerator: "CmdOrCtrl+S",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.saveWorkbook);
            });
          },
        },
        {
          label: "Save Workbook As...",
          accelerator: "CmdOrCtrl+Shift+S",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.saveWorkbookAs);
            });
          },
        },
        { type: "separator" },
        {
          label: "Import CSV...",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.importCsv);
            });
          },
        },
        {
          label: "Export CSV...",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.exportCsv);
            });
          },
        },
        { type: "separator" },
        {
          label: "Installation...",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.installation);
            });
          },
        },
        {
          label: "Check for Updates...",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.checkUpdates);
            });
          },
        },
        { type: "separator" },
        {
          label: "Exit",
          accelerator: "Alt+F4",
          click: () => {
            runMenuCommand(() => {
              app.quit();
            });
          },
        },
      ],
    },
    {
      enabled: menuEnabled,
      label: "Edit",
      submenu: [
        {
          accelerator: "CmdOrCtrl+Z",
          label: "Undo",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.undo);
            });
          },
        },
        {
          accelerator: "CmdOrCtrl+Shift+Z",
          label: "Redo",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.redo);
            });
          },
        },
        { type: "separator" },
        {
          accelerator: "CmdOrCtrl+X",
          label: "Cut",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.cut);
            });
          },
        },
        {
          accelerator: "CmdOrCtrl+Shift+X",
          label: "Cut Values",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.cutValues);
            });
          },
        },
        { type: "separator" },
        {
          accelerator: "CmdOrCtrl+C",
          label: "Copy",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.copy);
            });
          },
        },
        {
          accelerator: "CmdOrCtrl+Shift+C",
          label: "Copy Values",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.copyValues);
            });
          },
        },
        { type: "separator" },
        {
          accelerator: "CmdOrCtrl+V",
          label: "Paste",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.paste);
            });
          },
        },
        {
          accelerator: "CmdOrCtrl+Shift+V",
          label: "Paste Values",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.pasteValues);
            });
          },
        },
        { type: "separator" },
        {
          label: "Delete",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.deleteSelection);
            });
          },
        },
        { type: "separator" },
        {
          label: "Format...",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.formatCells);
            });
          },
        },
        {
          label: "Clear Formatting",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.clearFormatting);
            });
          },
        },
      ],
    },
    {
      enabled: menuEnabled,
      label: "Insert",
      submenu: [
        {
          label: "Table",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.insertTable);
            });
          },
        },
        {
          label: "Chart...",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.insertChart);
            });
          },
        },
      ],
    },
    {
      enabled: menuEnabled,
      label: "Sheet",
      submenu: [
        {
          label: "Add Row",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.addRow);
            });
          },
        },
        {
          label: "Add Column",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.addColumn);
            });
          },
        },
        { type: "separator" },
        {
          label: "New Sheet",
          accelerator: "CmdOrCtrl+Shift+N",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.newSheet);
            });
          },
        },
        {
          label: "Rename Sheet...",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.renameSheet);
            });
          },
        },
        {
          label: "Delete Sheet",
          click: () => {
            runMenuCommand(() => {
              sendMenuAction(APP_MENU_ACTIONS.deleteSheet);
            });
          },
        },
      ],
    },
    {
      enabled: menuEnabled,
      label: "Help",
      submenu: [
        {
          accelerator: "CmdOrCtrl+Shift+I",
          label: "Toggle Developer Tools",
          role: "toggleDevTools",
        },
        { type: "separator" },
        {
          label: "About...",
          click: () => {
            runMenuCommand(() => showAboutDialog());
          },
        },
      ],
    },
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

const createWindow = () => {
  startupTimer.log("create-window-start");
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
    icon: APP_ICON_PATH,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      sandbox: true,
    },
  });
  startupTimer.log("main-window-created", `windowId=${mainWindow.id}`);
  mainWindow.setTitle(formatWorkbookWindowTitle(workbookController.getSummary(), APP_DISPLAY_NAME));

  let didLogMainFrameFinishLoad = false;

  mainWindow.webContents.once("did-start-loading", () => {
    startupTimer.log("main-window-did-start-loading", `windowId=${mainWindow.id}`);
  });

  mainWindow.webContents.once("dom-ready", () => {
    startupTimer.log("main-window-dom-ready", `windowId=${mainWindow.id}`);
  });

  mainWindow.webContents.on("did-frame-finish-load", (_event, isMainFrame) => {
    if (!isMainFrame || didLogMainFrameFinishLoad) {
      return;
    }

    didLogMainFrameFinishLoad = true;
    startupTimer.log("main-window-main-frame-finish-load", `windowId=${mainWindow.id}`);
  });

  mainWindow.webContents.once("did-finish-load", () => {
    startupTimer.log("main-window-did-finish-load", `windowId=${mainWindow.id}`);
    flushPendingAppShowRequest(mainWindow, "did-finish-load");
  });

  mainWindow.webContents.once("did-stop-loading", () => {
    startupTimer.log("main-window-did-stop-loading", `windowId=${mainWindow.id}`);
    flushPendingAppShowRequest(mainWindow, "did-stop-loading");
  });

  mainWindow.webContents.once("did-fail-load", (_event, errorCode, errorDescription) => {
    startupTimer.log(
      "main-window-did-fail-load",
      `errorCode=${errorCode} errorDescription=${errorDescription}`,
    );
  });

  mainWindow.webContents.once("preload-error", (_event, preloadPath, error) => {
    startupTimer.log(
      "main-window-preload-error",
      `windowId=${mainWindow.id} preloadPath=${preloadPath} error=${error.message}`,
    );
  });

  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    startupTimer.log("main-window-load-url-start", MAIN_WINDOW_VITE_DEV_SERVER_URL);
    void mainWindow
      .loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL)
      .then(() => {
        startupTimer.log("main-window-load-url-done", MAIN_WINDOW_VITE_DEV_SERVER_URL);
      })
      .catch((error: unknown) => {
        startupTimer.log(
          "main-window-load-url-failed",
          error instanceof Error ? error.message : "unknown error",
        );
      });
    startupTimer.log("main-window-load-url-dispatched", MAIN_WINDOW_VITE_DEV_SERVER_URL);
  } else {
    const rendererPath = path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`);

    startupTimer.log("main-window-load-file-start", rendererPath);
    void mainWindow
      .loadFile(rendererPath)
      .then(() => {
        startupTimer.log("main-window-load-file-done", rendererPath);
      })
      .catch((error: unknown) => {
        startupTimer.log(
          "main-window-load-file-failed",
          error instanceof Error ? error.message : "unknown error",
        );
      });
    startupTimer.log("main-window-load-file-dispatched", rendererPath);
  }

  mainWindow.once("ready-to-show", () => {
    startupTimer.log("main-window-ready-to-show", `windowId=${mainWindow.id}`);
    if (isAppShowRequested) {
      flushPendingAppShowRequest(mainWindow, "ready-to-show");
      return;
    }

    if (!mainWindow.isVisible()) {
      mainWindow.show();
      startupTimer.log("main-window-show-called", `windowId=${mainWindow.id}`);
    }
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

  startupTimer.log("create-window-done");

  return mainWindow;
};

ipcMain.on("startup:timing", (event, message: StartupTimingMessage) => {
  const startupEvent = getStartupTimingMessageEvent(message);

  if (!startupEvent) {
    startupTimer.log("renderer-startup-timing-invalid");
    return;
  }

  startupTimer.log(
    `renderer-${startupEvent}`,
    formatStartupTimingMessageDetail(
      BrowserWindow.fromWebContents(event.sender),
      getStartupTimingMessageDetail(message),
    ),
  );
});

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
      error instanceof Error ? error.message : "The CSV file could not be opened.",
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

    const openResult = await openWorkbookFilePath(
      result.filePaths[0],
      browserWindow,
      unsavedChangesResolution,
    );

    if (!openResult) {
      return { canceled: true as const };
    }

    return {
      canceled: false as const,
      ...openResult,
    };
  } catch (error) {
    dialog.showErrorBox(
      "Open workbook failed",
      error instanceof Error ? error.message : "The workbook file could not be opened.",
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
      error instanceof Error ? error.message : "The CSV file could not be saved.",
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

ipcMain.handle("menu:show-cell-context-menu", async (event, args: ShowCellContextMenuArgs) => {
  const browserWindow = BrowserWindow.fromWebContents(event.sender);

  if (!browserWindow) {
    return;
  }

  buildCellContextMenu(browserWindow, args).popup({
    window: browserWindow,
  });
});

ipcMain.handle("menu:set-chart-dialog-open", (_event, isOpen: boolean) => {
  if (isChartDialogOpen === isOpen) {
    return;
  }

  isChartDialogOpen = isOpen;
  buildAppMenu();
});

ipcMain.handle("dialog:save-workbook-file-as", async (event, args?: SaveWorkbookFileAsArgs) => {
  const browserWindow = BrowserWindow.fromWebContents(event.sender);
  const filePath = await chooseWorkbookSavePath(browserWindow, args?.defaultPath);

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
});

ipcMain.handle("installer:get-status", () => installerService.getStatus());

ipcMain.handle("installer:install-current-app", (_event, options: InstallerOptions) =>
  installerService.installCurrentApp(options),
);

ipcMain.handle("installer:apply-options", (_event, options: InstallerOptions) =>
  installerService.applyOptions(options),
);

ipcMain.handle("installer:start-uninstall", () => installerService.startUninstall());

ipcMain.handle("installer:check-for-updates", (_event, request?: InstallerCheckUpdatesRequest) =>
  installerService.checkForUpdates(request),
);

ipcMain.handle("workbook:apply-transaction", (_event, args: ApplyTransactionRequest) =>
  workbookController.applyTransaction(args),
);

ipcMain.handle("workbook:get-undo-tree", () => workbookController.getUndoTree());

ipcMain.handle("workbook:redo", (_event, args?: WorkbookRedoRequest) =>
  workbookController.redo(args ?? {}),
);

ipcMain.handle("workbook:undo", (_event, args?: WorkbookHistoryRequest) =>
  workbookController.undo(args ?? {}),
);

ipcMain.handle("workbook:get-cell-data", (_event, args: CellDataRequest) =>
  workbookController.getCellData(args),
);

ipcMain.handle("workbook:get-chart", (_event, args: { chartId: string }) =>
  workbookController.getChart(args.chartId),
);

ipcMain.handle("workbook:get-chart-preview", (_event, args: { chartId: string }) =>
  workbookController.getChartPreview(args.chartId),
);

ipcMain.handle("workbook:copy-range", (_event, args: CopyRangeRequest) =>
  workbookController.copyRange(args),
);

ipcMain.handle("workbook:cut-range", (_event, args: CutRangeRequest) =>
  workbookController.cutRange(args),
);

ipcMain.handle("workbook:get-display-range", (_event, args: SheetRangeRequest) =>
  workbookController.getSheetDisplayRange(args),
);

ipcMain.handle("workbook:get-range", (_event, args: SheetRangeRequest) =>
  workbookController.getSheetRange(args),
);

ipcMain.handle("workbook:get-style-range", (_event, args: SheetRangeRequest) =>
  workbookController.getSheetStyleRange(args),
);

ipcMain.handle("workbook:paste-range", (_event, args: PasteRangeRequest) =>
  workbookController.pasteRange(args),
);

ipcMain.handle("workbook:get-sheet-csv", (_event, args?: { sheetId?: string }) =>
  workbookController.getSheetCsv(args?.sheetId),
);

ipcMain.handle("workbook:get-sheet-charts", (_event, args?: { sheetId?: string }) =>
  workbookController.getSheetCharts(args?.sheetId),
);

ipcMain.handle("workbook:get-sheet-chart-previews", (_event, args?: { sheetId?: string }) =>
  workbookController.getSheetChartPreviews(args?.sheetId),
);

ipcMain.handle("workbook:save-file", (_event, args: { filePath: string }) =>
  saveCurrentWorkbook(undefined, args.filePath).then((result) => {
    if (!result) {
      throw new Error("The workbook file could not be saved.");
    }

    return result;
  }),
);

ipcMain.handle("workbook:get-summary", () => workbookController.getSummary());

ipcMain.handle("workbook:get-used-range", (_event, args?: { sheetId?: string }) =>
  workbookController.getUsedRange(args?.sheetId),
);

workbookController.on("changed", () => {
  if (isConsoleExitMode) {
    return;
  }

  broadcastWorkbookChanged();
});

async function runConsoleOutputMode(filePath: string) {
  await workbookController.openWorkbookFile({
    discardUnsavedChanges: true,
    filePath,
  });
  process.stdout.write(workbookController.getConsoleOutput().text);
}

async function openStartupWorkbookFile(filePath: string | undefined) {
  if (!filePath) {
    return;
  }

  try {
    await workbookController.openWorkbookFile({
      discardUnsavedChanges: true,
      filePath,
    });
    startupTimer.log("startup-workbook-opened", filePath);
  } catch (error) {
    startupTimer.log(
      "startup-workbook-open-failed",
      error instanceof Error ? error.message : "unknown error",
    );
    dialog.showErrorBox(
      "Open workbook failed",
      error instanceof Error ? error.message : "The workbook file could not be opened.",
    );
  }
}

if (mainStartupOptions.help) {
  console.log(getMainHelpText(process.argv[0] ?? "spready"));
  app.exit(0);
} else if (mainStartupOptions.consoleOutputFilePath) {
  runConsoleOutputMode(mainStartupOptions.consoleOutputFilePath)
    .then(() => {
      app.exit(0);
    })
    .catch((error: unknown) => {
      console.error(error instanceof Error ? error.message : "Console output failed.");
      app.exit(1);
    });
} else if (isInstallerUninstallCommand) {
  app
    .whenReady()
    .then(() => installerService.startUninstall())
    .catch((error) => {
      dialog.showErrorBox(
        "Uninstall failed",
        error instanceof Error ? error.message : "Spready could not be uninstalled.",
      );
      app.quit();
    });
} else {
  app.whenReady().then(async () => {
    startupTimer.log("app-when-ready");
    startupTimer.log("recent-workbooks-load-start", recentWorkbooksStore.filePath);
    await recentWorkbooksStore
      .load()
      .then((result) => {
        startupTimer.log("recent-workbooks-load-done", `count=${result.workbooks.length}`);
      })
      .catch((error) => {
        startupTimer.log(
          "recent-workbooks-load-failed",
          error instanceof Error ? error.message : "unknown error",
        );
      });
    await openStartupWorkbookFile(mainStartupOptions.workbookFilePath);
    startupTimer.log("control-server-start-requested");
    void controlServer
      .start()
      .then(() => {
        const controlInfo = controlServer.getInfo();
        startupTimer.log("control-server-started", `tcp://${controlInfo.host}:${controlInfo.port}`);
        startupTimer.log("control-discovery-write-start");
        void writeDiscoveredControlInfo(APP_DISPLAY_NAME, controlInfo)
          .then(() => {
            startupTimer.log("control-discovery-write-done");
          })
          .catch((error) => {
            startupTimer.log(
              "control-discovery-write-failed",
              error instanceof Error ? error.message : "unknown error",
            );
          });
        console.log(
          `${APP_DISPLAY_NAME} control server listening on tcp://${controlInfo.host}:${controlInfo.port}`,
        );
      })
      .catch((error) => {
        startupTimer.log(
          "control-server-start-failed",
          error instanceof Error ? error.message : "unknown error",
        );
        console.error(
          `${APP_DISPLAY_NAME} control server failed to start: ${
            error instanceof Error ? error.message : "unknown error"
          }`,
        );
      });

    createWindow();
    buildAppMenu();
    startupTimer.log("initial-window-and-menu-created");

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      }
    });
  });
}

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("before-quit", () => {
  void clearDiscoveredControlInfo().catch((error) => {
    startupTimer.log(
      "control-discovery-clear-failed",
      error instanceof Error ? error.message : "unknown error",
    );
  });
  void controlServer.stop().catch((error) => {
    startupTimer.log(
      "control-server-stop-failed",
      error instanceof Error ? error.message : "unknown error",
    );
  });
});
