import net, { type Socket } from "node:net";

import type { StartupTimingLogger } from "./startup-timing";
import type { WorkbookController } from "./workbook-controller";
import type {
  AddRecentWorkbookRequest,
  ApplyTransactionRequest,
  ClearRangeRequest,
  ControlAppStatus,
  ControlServerInfo,
  CopyRangeRequest,
  CreateChartRequest,
  CreateNewWorkbookRequest,
  CutRangeRequest,
  ExportCsvFileRequest,
  FormatCellsRequest,
  ImportCsvFileRequest,
  OpenWorkbookFileRequest,
  SaveWorkbookFileRequest,
  PasteRangeRequest,
  RecentWorkbooksResult,
  RemoveRecentWorkbookRequest,
  SheetRangeRequest,
  WorkbookHistoryCheckoutRequest,
  WorkbookHistoryRequest,
  WorkbookRedoRequest,
  WorkbookSummary,
} from "./workbook-core";

type ControlRequest = {
  id?: number | string | null;
  method: string;
  params?: unknown;
};

type ControlErrorResponse = {
  error: string;
  id: number | string | null;
  ok: false;
};

type ControlSuccessResponse = {
  id: number | string | null;
  ok: true;
  result: unknown;
};

type ControlEvent = {
  event: string;
  payload: unknown;
};

const CONTROL_PROTOCOL = "spready-control-v1";

type SpreadyControlServerOptions = {
  addRecentWorkbook?: (
    request: AddRecentWorkbookRequest,
  ) => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  clearRecentWorkbooks?: () => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  getAppStatus?: () => ControlAppStatus;
  getRecentWorkbooks?: () => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  removeRecentWorkbook?: (
    request: RemoveRecentWorkbookRequest,
  ) => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  showApp?: () => ControlAppStatus | Promise<ControlAppStatus | void> | void;
  startupTimer?: StartupTimingLogger;
};

const DEFAULT_APP_STATUS: ControlAppStatus = {
  focusedWindowCount: 0,
  frontendVisible: false,
  visibleWindowCount: 0,
  windowCount: 0,
};

function shouldLogControlRequest(method: string, durationMs: number) {
  return (
    ["getAppStatus", "getControlInfo", "ping", "showApp"].includes(method) || durationMs >= 1000
  );
}

function isServerNotRunningError(error: Error | undefined): boolean {
  return (error as NodeJS.ErrnoException | undefined)?.code === "ERR_SERVER_NOT_RUNNING";
}

export class SpreadyControlServer {
  #addRecentWorkbook?: (
    request: AddRecentWorkbookRequest,
  ) => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  #clients = new Set<Socket>();
  #clearRecentWorkbooks?: () => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  #controller: WorkbookController;
  #getAppStatus: () => ControlAppStatus;
  #getRecentWorkbooks?: () => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  #host: string;
  #port: number;
  #removeRecentWorkbook?: (
    request: RemoveRecentWorkbookRequest,
  ) => Promise<RecentWorkbooksResult> | RecentWorkbooksResult;
  #server?: net.Server;
  #showApp?: () => ControlAppStatus | Promise<ControlAppStatus | void> | void;
  #startupTimer?: StartupTimingLogger;

  constructor(
    controller: WorkbookController,
    host: string,
    port: number,
    options: SpreadyControlServerOptions = {},
  ) {
    this.#addRecentWorkbook = options.addRecentWorkbook;
    this.#clearRecentWorkbooks = options.clearRecentWorkbooks;
    this.#controller = controller;
    this.#getAppStatus = options.getAppStatus ?? (() => DEFAULT_APP_STATUS);
    this.#getRecentWorkbooks = options.getRecentWorkbooks;
    this.#host = host;
    this.#port = port;
    this.#removeRecentWorkbook = options.removeRecentWorkbook;
    this.#showApp = options.showApp;
    this.#startupTimer = options.startupTimer;
  }

  getInfo(): ControlServerInfo {
    const address = this.#server?.address();
    const activePort =
      typeof address === "object" && address && "port" in address ? address.port : this.#port;

    return {
      host: this.#host,
      port: activePort,
      protocol: "jsonl",
    };
  }

  async start() {
    try {
      this.#logStartup("tcp-listen-start", `host=${this.#host} port=${this.#port}`);
      await this.#listen(this.#port);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EADDRINUSE") {
        throw error;
      }

      this.#logStartup("tcp-listen-port-in-use", `host=${this.#host} port=${this.#port}`);
      this.#logStartup("tcp-listen-start", `host=${this.#host} port=0`);
      await this.#listen(0);
    }

    const info = this.getInfo();
    this.#logStartup("tcp-listen-done", `tcp://${info.host}:${info.port}`);
    this.#controller.on("changed", this.#handleWorkbookChanged);
  }

  async stop() {
    this.#controller.off("changed", this.#handleWorkbookChanged);

    for (const socket of this.#clients) {
      socket.destroy();
    }

    this.#clients.clear();

    const server = this.#server;

    this.#server = undefined;

    if (!server) {
      return;
    }

    await new Promise<void>((resolve, reject) => {
      server.close((error) => {
        if (isServerNotRunningError(error)) {
          resolve();
          return;
        }

        if (error) {
          reject(error);
          return;
        }

        resolve();
      });
    });
  }

  #handleConnection = (socket: Socket) => {
    this.#clients.add(socket);
    socket.setEncoding("utf8");
    this.#logStartup("tcp-client-connected", `clients=${this.#clients.size}`);

    let buffer = "";

    this.#writeMessage(socket, {
      event: "hello",
      payload: {
        control: this.getInfo(),
        protocol: CONTROL_PROTOCOL,
        summary: this.#controller.getSummary(),
      },
    } satisfies ControlEvent);

    socket.on("data", (chunk: string) => {
      buffer += chunk;

      let newlineIndex = buffer.indexOf("\n");

      while (newlineIndex >= 0) {
        const line = buffer.slice(0, newlineIndex).trim();
        buffer = buffer.slice(newlineIndex + 1);

        if (line.length > 0) {
          void this.#handleLine(socket, line);
        }

        newlineIndex = buffer.indexOf("\n");
      }
    });

    socket.on("close", () => {
      this.#clients.delete(socket);
      this.#logStartup("tcp-client-closed", `clients=${this.#clients.size}`);
    });

    socket.on("error", () => {
      this.#clients.delete(socket);
      this.#logStartup("tcp-client-error", `clients=${this.#clients.size}`);
    });
  };

  async #handleLine(socket: Socket, line: string) {
    let request: ControlRequest;
    const requestStartedAt = Date.now();

    try {
      request = JSON.parse(line) as ControlRequest;
    } catch {
      this.#writeMessage(socket, {
        error: "Request must be valid JSON.",
        id: null,
        ok: false,
      } satisfies ControlErrorResponse);
      return;
    }

    if (typeof request.method !== "string" || request.method.length === 0) {
      this.#writeMessage(socket, {
        error: "Request must include a method string.",
        id: request.id ?? null,
        ok: false,
      } satisfies ControlErrorResponse);
      return;
    }

    try {
      const result = await this.#dispatchRequest(request.method, request.params);
      const durationMs = Date.now() - requestStartedAt;

      if (shouldLogControlRequest(request.method, durationMs)) {
        this.#logStartup("tcp-request-done", `method=${request.method} durationMs=${durationMs}`);
      }

      this.#writeMessage(socket, {
        id: request.id ?? null,
        ok: true,
        result,
      } satisfies ControlSuccessResponse);
    } catch (error) {
      const durationMs = Date.now() - requestStartedAt;

      if (shouldLogControlRequest(request.method, durationMs)) {
        this.#logStartup(
          "tcp-request-failed",
          `method=${request.method} durationMs=${durationMs} error=${
            error instanceof Error ? error.message : "Request failed."
          }`,
        );
      }

      this.#writeMessage(socket, {
        error: error instanceof Error ? error.message : "Request failed.",
        id: request.id ?? null,
        ok: false,
      } satisfies ControlErrorResponse);
    }
  }

  async #listen(port: number) {
    await new Promise<void>((resolve, reject) => {
      const server = net.createServer(this.#handleConnection);

      server.once("error", reject);
      server.listen(port, this.#host, () => {
        server.off("error", reject);
        this.#server = server;
        resolve();
      });
    });
  }

  #writeMessage(
    socket: Socket,
    message: ControlErrorResponse | ControlEvent | ControlSuccessResponse,
  ) {
    if (socket.destroyed) {
      return;
    }

    socket.write(`${JSON.stringify(message)}\n`);
  }

  #broadcast(event: ControlEvent) {
    for (const socket of this.#clients) {
      this.#writeMessage(socket, event);
    }
  }

  #handleWorkbookChanged = (summary: WorkbookSummary) => {
    this.#broadcast({
      event: "workbookChanged",
      payload: summary,
    });
  };

  #logStartup(event: string, detail?: string) {
    this.#startupTimer?.log(event, detail);
  }

  async #dispatchRequest(method: string, params: unknown) {
    switch (method) {
      case "applyTransaction":
        return this.#controller.applyTransaction(params as ApplyTransactionRequest);
      case "checkoutUndoNode":
        return this.#controller.checkoutUndoNode(params as WorkbookHistoryCheckoutRequest);
      case "clearRange":
        return this.#controller.clearRange(params as ClearRangeRequest);
      case "copyRange":
        return this.#controller.copyRange(params as CopyRangeRequest);
      case "cutRange":
        return this.#controller.cutRange(params as CutRangeRequest);
      case "createNewWorkbook":
        return this.#controller.createNewWorkbook(
          (params as CreateNewWorkbookRequest | undefined) ?? {},
        );
      case "createChart":
        return this.#controller.createChart(params as CreateChartRequest);
      case "exportCsvFile":
        return this.#controller.exportCsvFile(params as ExportCsvFileRequest);
      case "formatCells":
        return this.#controller.formatCells(params as FormatCellsRequest);
      case "getCellData":
        return this.#controller.getCellData(
          params as { columnIndex: number; rowIndex: number; sheetId?: string },
        );
      case "getChart":
        return this.#controller.getChart((params as { chartId: string }).chartId);
      case "getChartPreview":
        return this.#controller.getChartPreview((params as { chartId: string }).chartId);
      case "getControlInfo":
        return this.getInfo();
      case "getAppStatus":
        return this.#getAppStatus();
      case "getConsoleOutput":
        return this.#controller.getConsoleOutput();
      case "getSheetCsv":
        return this.#controller.getSheetCsv((params as { sheetId?: string } | undefined)?.sheetId);
      case "getRecentWorkbooks":
        return this.#getRecentWorkbooks?.() ?? { filePath: "", workbooks: [] };
      case "getSheetCharts":
        return this.#controller.getSheetCharts(
          (params as { sheetId?: string } | undefined)?.sheetId,
        );
      case "getSheetChartPreviews":
        return this.#controller.getSheetChartPreviews(
          (params as { sheetId?: string } | undefined)?.sheetId,
        );
      case "getSheetDisplayRange":
        return this.#controller.getSheetDisplayRange(params as SheetRangeRequest);
      case "getSheetRange":
        return this.#controller.getSheetRange(params as SheetRangeRequest);
      case "getSheetStyleRange":
        return this.#controller.getSheetStyleRange(params as SheetRangeRequest);
      case "getSheetTables":
        return this.#controller.getSheetTables(
          (params as { sheetId?: string } | undefined)?.sheetId,
        );
      case "getTable":
        return this.#controller.getTable((params as { tableId: string }).tableId);
      case "getUsedRange":
        return this.#controller.getUsedRange((params as { sheetId?: string } | undefined)?.sheetId);
      case "getUndoTree":
        return this.#controller.getUndoTree();
      case "getWorkbookSummary":
        return this.#controller.getSummary();
      case "importCsvFile":
        return this.#controller.importCsvFile(params as ImportCsvFileRequest);
      case "addRecentWorkbook":
        return (
          this.#addRecentWorkbook?.(params as AddRecentWorkbookRequest) ?? {
            filePath: "",
            workbooks: [],
          }
        );
      case "clearRecentWorkbooks":
        return this.#clearRecentWorkbooks?.() ?? { filePath: "", workbooks: [] };
      case "openWorkbookFile": {
        const result = await this.#controller.openWorkbookFile(params as OpenWorkbookFileRequest);

        await this.#addRecentWorkbook?.({ filePath: result.filePath });

        return result;
      }
      case "pasteRange":
        return this.#controller.pasteRange(params as PasteRangeRequest);
      case "removeRecentWorkbook":
        return (
          this.#removeRecentWorkbook?.(params as RemoveRecentWorkbookRequest) ?? {
            filePath: "",
            workbooks: [],
          }
        );
      case "redo":
        return this.#controller.redo((params as WorkbookRedoRequest | undefined) ?? {});
      case "saveWorkbookFile": {
        const result = await this.#controller.saveWorkbookFile(params as SaveWorkbookFileRequest);

        await this.#addRecentWorkbook?.({ filePath: result.filePath });

        return result;
      }
      case "showApp": {
        await this.#showApp?.();

        return this.#getAppStatus();
      }
      case "undo":
        return this.#controller.undo((params as WorkbookHistoryRequest | undefined) ?? {});
      case "listMethods":
        return [
          "addRecentWorkbook",
          "applyTransaction",
          "checkoutUndoNode",
          "clearRange",
          "clearRecentWorkbooks",
          "copyRange",
          "cutRange",
          "createChart",
          "createNewWorkbook",
          "exportCsvFile",
          "formatCells",
          "getCellData",
          "getChart",
          "getChartPreview",
          "getAppStatus",
          "getConsoleOutput",
          "getControlInfo",
          "getRecentWorkbooks",
          "getSheetCsv",
          "getSheetCharts",
          "getSheetChartPreviews",
          "getSheetDisplayRange",
          "getSheetRange",
          "getSheetStyleRange",
          "getSheetTables",
          "getTable",
          "getUndoTree",
          "getUsedRange",
          "getWorkbookSummary",
          "importCsvFile",
          "listMethods",
          "openWorkbookFile",
          "pasteRange",
          "ping",
          "redo",
          "removeRecentWorkbook",
          "saveWorkbookFile",
          "showApp",
          "undo",
        ];
      case "ping":
        return {
          control: this.getInfo(),
          protocol: CONTROL_PROTOCOL,
        };
      default:
        throw new Error(`Unknown control method "${method}".`);
    }
  }
}
