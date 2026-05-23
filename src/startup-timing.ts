import { appendFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

export type StartupLogSink = (message: string) => void;

export type StartupTimingClock = () => number;

export type StartupTimingLogger = {
  log: (event: string, detail?: string) => void;
};

export const STARTUP_TIMING_LOG_FILE_PATH = path.join(os.tmpdir(), "spready-startup.log");

export function createStartupLogSink(
  consoleSink: StartupLogSink = console.error,
  logFilePath = STARTUP_TIMING_LOG_FILE_PATH,
): StartupLogSink {
  return (message) => {
    consoleSink(message);
    void appendFile(logFilePath, `${new Date().toISOString()} ${message}\n`, "utf8").catch(
      () => undefined,
    );
  };
}

export function formatStartupTimingLog(
  scope: string,
  event: string,
  elapsedMs: number,
  totalMs: number,
  detail?: string,
) {
  const detailSuffix = detail ? ` ${detail}` : "";

  return `[${scope}] +${Math.round(elapsedMs)}ms total=${Math.round(totalMs)}ms ${event}${detailSuffix}`;
}

export class StartupTimer {
  #clock: StartupTimingClock;
  #lastMs: number;
  #scope: string;
  #sink: StartupLogSink;
  #startedMs: number;

  constructor(
    scope: string,
    sink: StartupLogSink = console.error,
    clock: StartupTimingClock = Date.now,
  ) {
    this.#clock = clock;
    this.#lastMs = this.#clock();
    this.#scope = scope;
    this.#sink = sink;
    this.#startedMs = this.#lastMs;
  }

  log(event: string, detail?: string) {
    const now = this.#clock();

    this.#sink(
      formatStartupTimingLog(this.#scope, event, now - this.#lastMs, now - this.#startedMs, detail),
    );
    this.#lastMs = now;
  }
}
