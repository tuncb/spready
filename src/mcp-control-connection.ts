import { EventEmitter } from "node:events";

import { openAppAndWaitForControlTarget, type McpStartupOptions } from "./mcp-startup";
import { resolveControlTarget, SpreadyControlClient, type ControlTarget } from "./control-client";
import { formatControlConnectionError } from "./mcp-control-errors";
import type { WorkbookSummary } from "./workbook-core";

type ConnectionState = "connected" | "connecting" | "disconnected" | "launching";

type ConnectionResult = {
  launched: boolean;
  target: ControlTarget;
};

type ConnectionStatus = {
  connected: boolean;
  lastError?: string;
  state: ConnectionState;
  target?: ControlTarget;
};

type ControlConnectionEventMap = {
  workbookChanged: WorkbookSummary;
};

type ControlConnectionDependencies = {
  createClient?: (target: ControlTarget) => SpreadyControlClient;
  openAppAndWaitForControlTarget?: typeof openAppAndWaitForControlTarget;
  resolveControlTarget?: typeof resolveControlTarget;
};

export class McpControlConnection extends EventEmitter {
  #client?: SpreadyControlClient;
  #createClient: (target: ControlTarget) => SpreadyControlClient;
  #lastError?: string;
  #openAppAndWaitForControlTarget: typeof openAppAndWaitForControlTarget;
  #operation?: Promise<ConnectionResult>;
  #resolveControlTarget: typeof resolveControlTarget;
  #startupOptions: McpStartupOptions;
  #state: ConnectionState = "disconnected";
  #target?: ControlTarget;

  constructor(startupOptions: McpStartupOptions, dependencies: ControlConnectionDependencies = {}) {
    super();
    this.#startupOptions = startupOptions;
    this.#createClient =
      dependencies.createClient ?? ((target) => new SpreadyControlClient(target));
    this.#openAppAndWaitForControlTarget =
      dependencies.openAppAndWaitForControlTarget ?? openAppAndWaitForControlTarget;
    this.#resolveControlTarget = dependencies.resolveControlTarget ?? resolveControlTarget;
  }

  getStatus(): ConnectionStatus {
    const status: ConnectionStatus = {
      connected: this.#state === "connected" && this.#client !== undefined,
      state: this.#state,
    };

    if (this.#lastError) {
      status.lastError = this.#lastError;
    }

    if (this.#target) {
      status.target = this.#target;
    }

    return status;
  }

  requireConnectedClient() {
    if (!this.#client || this.#state !== "connected") {
      throw new Error(
        "Spready app is not connected. Call open_spready_app first, or start Spready and call open_spready_app to connect.",
      );
    }

    return this.#client;
  }

  async connectToExisting(): Promise<ConnectionResult> {
    if (this.#client && this.#target && this.#state === "connected") {
      return {
        launched: false,
        target: this.#target,
      };
    }

    return this.#runExclusive("connecting", async () => {
      const target = await this.#resolveControlTarget({
        host: this.#startupOptions.host,
        port: this.#startupOptions.port,
      });

      await this.#connectClient(target);

      return {
        launched: false,
        target,
      };
    });
  }

  async openAppAndConnect(): Promise<ConnectionResult> {
    if (this.#client && this.#target && this.#state === "connected") {
      return {
        launched: false,
        target: this.#target,
      };
    }

    try {
      return await this.connectToExisting();
    } catch {
      // Fall through to launching a fresh app. The launch failure, if any, becomes public.
    }

    return this.launchAppAndConnect();
  }

  async launchAppAndConnect(): Promise<ConnectionResult> {
    if (this.#client && this.#target && this.#state === "connected") {
      return {
        launched: false,
        target: this.#target,
      };
    }

    return this.#runExclusive("launching", async () => {
      const target = await this.#openAppAndWaitForControlTarget(this.#startupOptions);

      await this.#connectClient(target);

      return {
        launched: true,
        target,
      };
    });
  }

  override on<EventName extends keyof ControlConnectionEventMap>(
    eventName: EventName,
    listener: (payload: ControlConnectionEventMap[EventName]) => void,
  ): this {
    return super.on(eventName, listener);
  }

  async #connectClient(target: ControlTarget) {
    const client = this.#createClient(target);

    try {
      await client.connect();
    } catch (error) {
      throw new Error(formatControlConnectionError(target, error));
    }

    this.#client = client;
    this.#target = target;
    this.#state = "connected";
    this.#lastError = undefined;

    client.on("workbookChanged", (summary) => {
      if (this.#client === client) {
        this.emit("workbookChanged", summary);
      }
    });

    (client as EventEmitter).on("close", () => {
      if (this.#client !== client) {
        return;
      }

      this.#client = undefined;
      this.#state = "disconnected";
    });
  }

  #formatError(error: unknown) {
    return error instanceof Error ? error.message : "unknown connection error";
  }

  async #runExclusive(
    state: Exclude<ConnectionState, "connected" | "disconnected">,
    operation: () => Promise<ConnectionResult>,
  ) {
    if (this.#operation) {
      return this.#operation;
    }

    this.#state = state;
    this.#lastError = undefined;
    this.#operation = operation()
      .then((result) => {
        this.#state = "connected";
        this.#target = result.target;

        return result;
      })
      .catch((error: unknown) => {
        this.#client = undefined;
        this.#state = "disconnected";
        this.#lastError = this.#formatError(error);
        throw error;
      })
      .finally(() => {
        this.#operation = undefined;
      });

    return this.#operation;
  }
}
