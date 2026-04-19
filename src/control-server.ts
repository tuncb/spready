import net, { type Socket } from "node:net";

import type { WorkbookController } from "./workbook-controller";
import type {
  ApplyTransactionRequest,
  ClearRangeRequest,
  ControlServerInfo,
  CopyRangeRequest,
  ExportCsvFileRequest,
  ImportCsvFileRequest,
  PasteRangeRequest,
  SheetRangeRequest,
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

export class SpreadyControlServer {
  #clients = new Set<Socket>();
  #controller: WorkbookController;
  #host: string;
  #port: number;
  #server?: net.Server;

  constructor(controller: WorkbookController, host: string, port: number) {
    this.#controller = controller;
    this.#host = host;
    this.#port = port;
  }

  getInfo(): ControlServerInfo {
    const address = this.#server?.address();
    const activePort =
      typeof address === "object" && address && "port" in address
        ? address.port
        : this.#port;

    return {
      host: this.#host,
      port: activePort,
      protocol: "jsonl",
    };
  }

  async start() {
    try {
      await this.#listen(this.#port);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EADDRINUSE") {
        throw error;
      }

      await this.#listen(0);
    }

    this.#controller.on("changed", this.#handleWorkbookChanged);
  }

  async stop() {
    this.#controller.off("changed", this.#handleWorkbookChanged);

    for (const socket of this.#clients) {
      socket.destroy();
    }

    this.#clients.clear();

    if (!this.#server) {
      return;
    }

    await new Promise<void>((resolve, reject) => {
      this.#server?.close((error) => {
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
    });

    socket.on("error", () => {
      this.#clients.delete(socket);
    });
  };

  async #handleLine(socket: Socket, line: string) {
    let request: ControlRequest;

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
      const result = await this.#dispatchRequest(
        request.method,
        request.params,
      );

      this.#writeMessage(socket, {
        id: request.id ?? null,
        ok: true,
        result,
      } satisfies ControlSuccessResponse);
    } catch (error) {
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

  async #dispatchRequest(method: string, params: unknown) {
    switch (method) {
      case "applyTransaction":
        return this.#controller.applyTransaction(
          params as ApplyTransactionRequest,
        );
      case "clearRange":
        return this.#controller.clearRange(params as ClearRangeRequest);
      case "copyRange":
        return this.#controller.copyRange(params as CopyRangeRequest);
      case "exportCsvFile":
        return this.#controller.exportCsvFile(params as ExportCsvFileRequest);
      case "getCellData":
        return this.#controller.getCellData(
          params as { columnIndex: number; rowIndex: number; sheetId?: string },
        );
      case "getControlInfo":
        return this.getInfo();
      case "getSheetCsv":
        return this.#controller.getSheetCsv(
          (params as { sheetId?: string } | undefined)?.sheetId,
        );
      case "getSheetDisplayRange":
        return this.#controller.getSheetDisplayRange(
          params as SheetRangeRequest,
        );
      case "getSheetRange":
        return this.#controller.getSheetRange(params as SheetRangeRequest);
      case "getUsedRange":
        return this.#controller.getUsedRange(
          (params as { sheetId?: string } | undefined)?.sheetId,
        );
      case "getWorkbookSummary":
        return this.#controller.getSummary();
      case "importCsvFile":
        return this.#controller.importCsvFile(params as ImportCsvFileRequest);
      case "pasteRange":
        return this.#controller.pasteRange(params as PasteRangeRequest);
      case "listMethods":
        return [
          "applyTransaction",
          "clearRange",
          "copyRange",
          "exportCsvFile",
          "getCellData",
          "getControlInfo",
          "getSheetCsv",
          "getSheetDisplayRange",
          "getSheetRange",
          "getUsedRange",
          "getWorkbookSummary",
          "importCsvFile",
          "listMethods",
          "pasteRange",
          "ping",
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
