import { spawn, type ChildProcess } from "node:child_process";
import { access, readFile } from "node:fs/promises";
import path from "node:path";

import { readDiscoveredControlInfo } from "./control-discovery";
import { resolveControlTarget, SpreadyControlClient, type ControlTarget } from "./control-client";
import { formatControlConnectionError } from "./mcp-control-errors";

const APP_DISPLAY_NAME = "Spready";
const DEFAULT_OPEN_APP_TIMEOUT_MS = 20000;
const DEFAULT_POLL_INTERVAL_MS = 200;

export interface McpStartupOptions {
  appPath?: string;
  help?: boolean;
  host?: string;
  openApp: boolean;
  openAppTimeoutMs: number;
  port?: number;
}

export interface AppLaunchCommand {
  args: string[];
  command: string;
  cwd?: string;
  env: NodeJS.ProcessEnv;
}

type ResolveAppLaunchOptions = {
  appPath?: string;
  cwd?: string;
  env?: NodeJS.ProcessEnv;
  executablePath?: string;
  host?: string;
  platform?: NodeJS.Platform;
  port?: number;
};

type WaitForControlTargetOptions = {
  host?: string;
  launchedAfter?: Date;
  pollIntervalMs?: number;
  port?: number;
  preferFreshDiscovery?: boolean;
  timeoutMs: number;
};

function getArgumentName(token: string) {
  const separatorIndex = token.indexOf("=");

  return token.slice(2, separatorIndex === -1 ? undefined : separatorIndex);
}

function getInlineArgumentValue(token: string) {
  const separatorIndex = token.indexOf("=");

  if (separatorIndex === -1) {
    return undefined;
  }

  return token.slice(separatorIndex + 1);
}

function readArgumentValue(argv: string[], index: number, token: string) {
  const inlineValue = getInlineArgumentValue(token);

  if (inlineValue !== undefined) {
    if (inlineValue.length === 0) {
      throw new Error(`Missing value for --${getArgumentName(token)}.`);
    }

    return {
      nextIndex: index,
      value: inlineValue,
    };
  }

  const value = argv[index + 1];

  if (!value || value.startsWith("--")) {
    throw new Error(`Missing value for --${getArgumentName(token)}.`);
  }

  return {
    nextIndex: index + 1,
    value,
  };
}

function parseIntegerArgument(value: string, name: string) {
  const parsed = Number.parseInt(value, 10);

  if (!Number.isInteger(parsed)) {
    throw new Error(`Invalid ${name} "${value}".`);
  }

  return parsed;
}

function parsePort(portValue: string) {
  const port = parseIntegerArgument(portValue, "port");

  if (port < 1 || port > 65535) {
    throw new Error(`Invalid port "${portValue}".`);
  }

  return port;
}

export function parseMcpStartupOptions(argv: string[]): McpStartupOptions {
  const options: McpStartupOptions = {
    help: false,
    openApp: false,
    openAppTimeoutMs: DEFAULT_OPEN_APP_TIMEOUT_MS,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];

    if (token === "-h") {
      options.help = true;
      continue;
    }

    if (!token.startsWith("--")) {
      continue;
    }

    const name = getArgumentName(token);

    switch (name) {
      case "help":
        if (getInlineArgumentValue(token) !== undefined) {
          throw new Error(`--${name} does not accept a value.`);
        }

        options.help = true;
        break;
      case "app-path":
      case "appPath": {
        const parsed = readArgumentValue(argv, index, token);
        options.appPath = parsed.value;
        index = parsed.nextIndex;
        break;
      }
      case "host": {
        const parsed = readArgumentValue(argv, index, token);
        options.host = parsed.value;
        index = parsed.nextIndex;
        break;
      }
      case "open-app":
      case "openApp":
        if (getInlineArgumentValue(token) !== undefined) {
          throw new Error(`--${name} does not accept a value.`);
        }

        options.openApp = true;
        break;
      case "open-app-timeout-ms":
      case "openAppTimeoutMs": {
        const parsed = readArgumentValue(argv, index, token);
        const timeoutMs = parseIntegerArgument(parsed.value, "open app timeout");

        if (timeoutMs < 1) {
          throw new Error(`Invalid open app timeout "${parsed.value}".`);
        }

        options.openAppTimeoutMs = timeoutMs;
        index = parsed.nextIndex;
        break;
      }
      case "port": {
        const parsed = readArgumentValue(argv, index, token);
        options.port = parsePort(parsed.value);
        index = parsed.nextIndex;
        break;
      }
      default:
        break;
    }
  }

  return options;
}

export function getMcpStdioHelpText(commandName = "spready-mcp") {
  return [
    "Spready MCP stdio wrapper",
    "",
    `Usage: ${commandName} [options]`,
    "",
    "Options:",
    "  -h, --help                      Show this help message.",
    "      --openApp, --open-app       Launch Spready during wrapper startup.",
    "      --appPath, --app-path PATH  Path to the Spready app executable or .app bundle.",
    "      --host HOST                 TCP control server host to connect to or launch with.",
    "      --port PORT                 TCP control server port to connect to or launch with.",
    "      --openAppTimeoutMs MS       Milliseconds to wait for a launched app control server.",
    "      --open-app-timeout-ms MS    Dashed alias for --openAppTimeoutMs.",
    "",
    "The wrapper speaks MCP over stdio. Without --openApp, it connects to an existing",
    "Spready control server when one is available; otherwise clients can call the",
    "open_spready_app MCP tool after startup.",
  ].join("\n");
}

export function getPackagedAppExecutablePath(appPath: string, platform = process.platform) {
  if (platform === "darwin" && appPath.endsWith(".app")) {
    return path.join(appPath, "Contents", "MacOS", APP_DISPLAY_NAME);
  }

  return appPath;
}

function getPackagedAppCandidates(baseDir: string, platform = process.platform) {
  if (platform === "win32") {
    return [path.join(baseDir, `${APP_DISPLAY_NAME}.exe`)];
  }

  if (platform === "darwin") {
    return [path.join(baseDir, `${APP_DISPLAY_NAME}.app`, "Contents", "MacOS", APP_DISPLAY_NAME)];
  }

  return [path.join(baseDir, APP_DISPLAY_NAME)];
}

async function pathExists(filePath: string) {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

function getNpmStartLaunchCommand(platform = process.platform) {
  if (platform === "win32") {
    return {
      args: ["/d", "/s", "/c", "npm.cmd run start"],
      command: process.env.ComSpec ?? "cmd.exe",
    };
  }

  return {
    args: ["run", "start"],
    command: "npm",
  };
}

async function isSpreadyRepository(cwd: string) {
  try {
    const packageJson = JSON.parse(await readFile(path.join(cwd, "package.json"), "utf8")) as {
      name?: unknown;
    };

    return packageJson.name === "spready";
  } catch {
    return false;
  }
}

function createLaunchEnvironment(
  options: Pick<McpStartupOptions, "host" | "port">,
  env: NodeJS.ProcessEnv,
) {
  const launchEnv: NodeJS.ProcessEnv = {
    ...env,
  };

  if (options.host) {
    launchEnv.SPREADY_CONTROL_HOST = options.host;
  }

  if (options.port !== undefined) {
    launchEnv.SPREADY_CONTROL_PORT = `${options.port}`;
  }

  return launchEnv;
}

export async function resolveSpreadyAppLaunchCommand(
  options: ResolveAppLaunchOptions = {},
): Promise<AppLaunchCommand> {
  const cwd = options.cwd ?? process.cwd();
  const platform = options.platform ?? process.platform;
  const env = createLaunchEnvironment(options, options.env ?? process.env);

  if (options.appPath) {
    const command = getPackagedAppExecutablePath(options.appPath, platform);

    if (!(await pathExists(command))) {
      throw new Error(`Spready app executable was not found at "${command}".`);
    }

    return {
      args: [],
      command,
      env,
    };
  }

  const executableDir = path.dirname(options.executablePath ?? process.execPath);
  const candidateDirs = Array.from(new Set([executableDir, cwd]));

  for (const candidateDir of candidateDirs) {
    for (const candidate of getPackagedAppCandidates(candidateDir, platform)) {
      if (await pathExists(candidate)) {
        return {
          args: [],
          command: candidate,
          env,
        };
      }
    }
  }

  if (await isSpreadyRepository(cwd)) {
    const npmStart = getNpmStartLaunchCommand(platform);

    return {
      args: npmStart.args,
      command: npmStart.command,
      cwd,
      env,
    };
  }

  throw new Error(
    `Could not find ${APP_DISPLAY_NAME}. Pass --appPath with the app executable path.`,
  );
}

export async function launchSpreadyApp(command: AppLaunchCommand): Promise<ChildProcess> {
  const child = spawn(command.command, command.args, {
    cwd: command.cwd,
    detached: true,
    env: command.env,
    stdio: "ignore",
    windowsHide: true,
  });

  await new Promise<void>((resolve, reject) => {
    const timer = setTimeout(resolve, 100);

    child.once("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
  });

  child.unref();

  return child;
}

async function connectAndClose(target: ControlTarget) {
  const client = new SpreadyControlClient(target);

  await client.connect();
  await client.close();
}

async function resolveFreshDiscoveryTarget(launchedAfter: Date, host?: string) {
  const discovered = await readDiscoveredControlInfo();

  if (!discovered) {
    return null;
  }

  const updatedAt = Date.parse(discovered.updatedAt);

  if (!Number.isFinite(updatedAt) || updatedAt < launchedAfter.getTime()) {
    return null;
  }

  return {
    host: host ?? discovered.host,
    port: discovered.port,
    source: "discovery",
  } satisfies ControlTarget;
}

async function resolveStartupTarget(options: WaitForControlTargetOptions) {
  if (options.preferFreshDiscovery) {
    return options.launchedAfter
      ? resolveFreshDiscoveryTarget(options.launchedAfter, options.host)
      : null;
  }

  return resolveControlTarget({
    host: options.host,
    port: options.port,
  });
}

function sleep(ms: number) {
  return new Promise<void>((resolve) => {
    setTimeout(resolve, ms);
  });
}

export async function waitForControlTarget(
  options: WaitForControlTargetOptions,
): Promise<ControlTarget> {
  const deadline = Date.now() + options.timeoutMs;
  let lastError: unknown;
  let lastTarget: ControlTarget | null = null;

  while (Date.now() <= deadline) {
    const target = await resolveStartupTarget(options);

    if (target) {
      lastTarget = target;
      try {
        await connectAndClose(target);
        return target;
      } catch (error) {
        lastError = error;
      }
    }

    await sleep(options.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS);
  }

  const detail =
    lastError instanceof Error && lastTarget
      ? ` ${formatControlConnectionError(lastTarget, lastError)}`
      : "";

  throw new Error(`Timed out waiting for the Spready control server.${detail}`);
}

export async function openAppAndWaitForControlTarget(options: McpStartupOptions) {
  const command = await resolveSpreadyAppLaunchCommand(options);
  const launchedAfter = new Date();

  await launchSpreadyApp(command);

  return waitForControlTarget({
    host: options.host,
    launchedAfter,
    port: options.port,
    preferFreshDiscovery: true,
    timeoutMs: options.openAppTimeoutMs,
  });
}
