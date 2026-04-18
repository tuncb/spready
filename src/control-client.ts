import { EventEmitter } from 'node:events';
import net, { type Socket } from 'node:net';

import { readDiscoveredControlInfo } from './control-discovery';
import type {
  ApplyTransactionRequest,
  ApplyTransactionResult,
  ControlServerInfo,
  SheetRangeRequest,
  SheetRangeResult,
  UsedRangeResult,
  WorkbookSummary,
} from './workbook-core';

const DEFAULT_CONTROL_HOST = '127.0.0.1';
const DEFAULT_CONTROL_PORT = 45731;
const DEFAULT_REQUEST_TIMEOUT_MS = 15000;

type ControlClientEventMap = {
  hello: {
    control: ControlServerInfo;
    protocol: string;
    summary: WorkbookSummary;
  };
  workbookChanged: WorkbookSummary;
};

type ControlRequest = {
  id: number;
  method: string;
  params?: unknown;
};

type ControlEventMessage = {
  event: keyof ControlClientEventMap | string;
  payload: unknown;
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

type PendingRequest = {
  reject: (error: Error) => void;
  resolve: (value: unknown) => void;
  timeout: NodeJS.Timeout;
};

export interface ControlTarget {
  host: string;
  port: number;
  source: 'argv' | 'default' | 'discovery' | 'env';
}

export interface ControlClientOptions {
  host?: string;
  port?: number;
}

export class SpreadyControlClient extends EventEmitter {
  #buffer = '';
  #host: string;
  #nextRequestId = 1;
  #pendingRequests = new Map<number, PendingRequest>();
  #port: number;
  #socket?: Socket;

  constructor(target: ControlTarget) {
    super();
    this.#host = target.host;
    this.#port = target.port;
  }

  async connect() {
    if (this.#socket && !this.#socket.destroyed) {
      return;
    }

    const socket = net.createConnection({
      host: this.#host,
      port: this.#port,
    });

    this.#socket = socket;
    socket.setEncoding('utf8');

    socket.on('data', (chunk: string) => {
      this.#buffer += chunk;

      let newlineIndex = this.#buffer.indexOf('\n');

      while (newlineIndex >= 0) {
        const line = this.#buffer.slice(0, newlineIndex).trim();
        this.#buffer = this.#buffer.slice(newlineIndex + 1);

        if (line.length > 0) {
          this.#handleMessage(line);
        }

        newlineIndex = this.#buffer.indexOf('\n');
      }
    });

    socket.on('error', (error) => {
      this.#rejectAllPending(error);
      this.emit('error', error);
    });

    socket.on('close', () => {
      this.#rejectAllPending(new Error('Spready control connection closed.'));
      this.emit('close');
    });

    await new Promise<void>((resolve, reject) => {
      socket.once('connect', resolve);
      socket.once('error', reject);
    });
  }

  async close() {
    if (!this.#socket || this.#socket.destroyed) {
      return;
    }

    await new Promise<void>((resolve) => {
      this.#socket?.end(() => resolve());
    });
  }

  async applyTransaction(request: ApplyTransactionRequest) {
    return this.call<ApplyTransactionResult>('applyTransaction', request);
  }

  async getSheetCsv(sheetId?: string) {
    return this.call<string>('getSheetCsv', { sheetId });
  }

  async getSheetRange(request: SheetRangeRequest) {
    return this.call<SheetRangeResult>('getSheetRange', request);
  }

  async getUsedRange(sheetId?: string) {
    return this.call<UsedRangeResult>('getUsedRange', { sheetId });
  }

  async getWorkbookSummary() {
    return this.call<WorkbookSummary>('getWorkbookSummary');
  }

  async call<Result>(method: string, params?: unknown): Promise<Result> {
    const socket = this.#socket;

    if (!socket || socket.destroyed) {
      throw new Error('Spready control client is not connected.');
    }

    const id = this.#nextRequestId;

    this.#nextRequestId += 1;

    const request: ControlRequest = {
      id,
      method,
      params,
    };

    return new Promise<Result>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.#pendingRequests.delete(id);
        reject(new Error(`Spready control request timed out for "${method}".`));
      }, DEFAULT_REQUEST_TIMEOUT_MS);

      this.#pendingRequests.set(id, {
        reject,
        resolve: (value) => resolve(value as Result),
        timeout,
      });

      socket.write(`${JSON.stringify(request)}\n`);
    });
  }

  override on<EventName extends keyof ControlClientEventMap>(
    eventName: EventName,
    listener: (payload: ControlClientEventMap[EventName]) => void,
  ): this {
    return super.on(eventName, listener);
  }

  #handleMessage(line: string) {
    const message = JSON.parse(line) as ControlErrorResponse | ControlEventMessage | ControlSuccessResponse;

    if ('event' in message) {
      this.emit(message.event, message.payload);
      return;
    }

    if (typeof message.id !== 'number') {
      return;
    }

    const pendingRequest = this.#pendingRequests.get(message.id);

    if (!pendingRequest) {
      return;
    }

    clearTimeout(pendingRequest.timeout);
    this.#pendingRequests.delete(message.id);

    if ('ok' in message && message.ok) {
      pendingRequest.resolve(message.result);
      return;
    }

    pendingRequest.reject(new Error(message.error));
  }

  #rejectAllPending(error: Error) {
    for (const pendingRequest of this.#pendingRequests.values()) {
      clearTimeout(pendingRequest.timeout);
      pendingRequest.reject(error);
    }

    this.#pendingRequests.clear();
  }
}

export async function resolveControlTarget(options: ControlClientOptions = {}): Promise<ControlTarget> {
  if (options.host || options.port) {
    return {
      host: options.host ?? DEFAULT_CONTROL_HOST,
      port: options.port ?? DEFAULT_CONTROL_PORT,
      source: 'argv',
    };
  }

  const envHost = process.env.SPREADY_CONTROL_HOST;
  const envPortValue = process.env.SPREADY_CONTROL_PORT;

  if (envHost || envPortValue) {
    const envPort = envPortValue ? Number.parseInt(envPortValue, 10) : DEFAULT_CONTROL_PORT;

    if (Number.isNaN(envPort)) {
      throw new Error('SPREADY_CONTROL_PORT must be a valid integer.');
    }

    return {
      host: envHost ?? DEFAULT_CONTROL_HOST,
      port: envPort,
      source: 'env',
    };
  }

  const discoveredTarget = await readDiscoveredControlInfo();

  if (discoveredTarget) {
    return {
      host: discoveredTarget.host,
      port: discoveredTarget.port,
      source: 'discovery',
    };
  }

  return {
    host: DEFAULT_CONTROL_HOST,
    port: DEFAULT_CONTROL_PORT,
    source: 'default',
  };
}
