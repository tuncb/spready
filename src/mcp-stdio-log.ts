import type { ControlTarget } from "./control-client";

export function formatDisconnectedStartupLog(discoveryFilePath: string, detail: string) {
  return [
    "Spready MCP stdio wrapper started without a control connection.",
    "Action: call open_spready_app to connect or launch Spready.",
    `Discovery file: ${discoveryFilePath}`,
    `Detail: ${detail}`,
  ].join("\n");
}

export function formatConnectedStartupLog(target: ControlTarget) {
  return [
    "Spready MCP stdio wrapper connected.",
    `Address: tcp://${target.host}:${target.port}`,
    `Source: ${target.source}`,
  ].join("\n");
}

export function formatReadyStartupLog() {
  return ["Spready MCP stdio wrapper ready.", "Action: call open_spready_app to connect."].join(
    "\n",
  );
}
