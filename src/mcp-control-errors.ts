import type { ControlTarget } from "./control-client";

function getErrorMessage(error: unknown) {
  return error instanceof Error ? error.message : "unknown connection error";
}

function getErrorCode(error: unknown) {
  return typeof error === "object" && error !== null && "code" in error
    ? (error as { code?: unknown }).code
    : undefined;
}

export function formatControlConnectionError(target: ControlTarget, error: unknown) {
  const address = `tcp://${target.host}:${target.port}`;
  const message = getErrorMessage(error);
  const code = getErrorCode(error);

  if (code === "ECONNREFUSED" || message.includes("ECONNREFUSED")) {
    return (
      `Could not connect to the Spready TCP control server at ${address}. ` +
      "Open Spready at this address through MCP or manually."
    );
  }

  return `Could not connect to the Spready TCP control server at ${address}. ${message}`;
}
