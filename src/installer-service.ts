import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

import type {
  InstallerCheckUpdatesRequest,
  InstallerCheckUpdatesResult,
  InstallerOperationResult,
  InstallerOptions,
  InstallerStatus,
} from "./workbook-core";

const APP_NAME = "Spready";
const APP_EXECUTABLE_NAME = "Spready.exe";
const FILE_ASSOCIATION_DESCRIPTION = "Spready Workbook";
const FILE_ASSOCIATION_EXTENSION = ".spready";
const FILE_ASSOCIATION_PROG_ID = "Spready.Workbook";
const GITHUB_LATEST_RELEASE_URL = "https://api.github.com/repos/tuncb/spready/releases/latest";
const PUBLISHER = "tuncb";
const FILE_ASSOCIATION_EXTENSION_KEY = `HKCU\\Software\\Classes\\${FILE_ASSOCIATION_EXTENSION}`;
const FILE_ASSOCIATION_PROG_ID_KEY = `HKCU\\Software\\Classes\\${FILE_ASSOCIATION_PROG_ID}`;
const UNINSTALL_REGISTRY_KEY =
  "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Spready";
const INSTALLER_LOG_DIRECTORY_NAME = "spready-installation-logs";
const MAX_RELEASE_RESPONSE_BYTES = 2 * 1024 * 1024;
const MAX_ASSET_DOWNLOAD_BYTES = 250 * 1024 * 1024;

export interface VersionNumber {
  major: number;
  minor: number;
  patch: number;
}

export interface ReleaseAsset {
  browser_download_url: string;
  digest?: string;
  name: string;
}

export interface LatestReleaseInfo {
  assets: ReleaseAsset[];
  html_url: string;
  tag_name: string;
  version: VersionNumber;
}

interface InstallerServiceDependencies {
  arch?: NodeJS.Architecture;
  commandRunner?: InstallerCommandRunner;
  currentAppDirectory: string;
  currentExecutablePath: string;
  currentVersion: string;
  detachedPowerShellRunner?: InstallerDetachedPowerShellRunner;
  env?: NodeJS.ProcessEnv;
  fetch?: typeof fetch;
  isPackaged: boolean;
  platform?: NodeJS.Platform;
  requestQuit?: () => void;
  writeShortcut?: (
    shortcutPath: string,
    operation: InstallerShortcutOperation,
    shortcut: InstallerShortcutDetails,
  ) => boolean;
}

export interface InstallerCommandResult {
  stderr: string;
  stdout: string;
}

export type InstallerCommandRunner = (
  command: string,
  args: string[],
) => Promise<InstallerCommandResult>;

export type InstallerDetachedPowerShellRunner = (
  operationName: string,
  script: string,
  options?: { logPath?: string },
) => Promise<void>;

export interface InstallerShortcutDetails {
  cwd: string;
  description: string;
  icon: string;
  iconIndex: number;
  target: string;
}

export type InstallerShortcutOperation = "create" | "update" | "replace";

export function parseVersionTag(tag: string): VersionNumber | null {
  const match = /^v?(\d+)\.(\d+)\.(\d+)(?:[-+].*)?$/u.exec(tag.trim());

  if (!match) {
    return null;
  }

  return {
    major: Number.parseInt(match[1], 10),
    minor: Number.parseInt(match[2], 10),
    patch: Number.parseInt(match[3], 10),
  };
}

export function formatVersionNumber(version: VersionNumber): string {
  return `${version.major}.${version.minor}.${version.patch}`;
}

export function isVersionNewer(candidate: VersionNumber, current: VersionNumber): boolean {
  if (candidate.major !== current.major) {
    return candidate.major > current.major;
  }

  if (candidate.minor !== current.minor) {
    return candidate.minor > current.minor;
  }

  return candidate.patch > current.patch;
}

export function extractSha256FromText(text: string): string | null {
  const match = /\b[a-fA-F0-9]{64}\b/u.exec(text);

  return match ? match[0].toLowerCase() : null;
}

export function extractSha256FromDigest(digest: string | undefined): string | null {
  if (!digest) {
    return null;
  }

  const normalizedDigest = digest.trim().toLowerCase();
  const prefixedMatch = /^sha256:([a-f0-9]{64})$/u.exec(normalizedDigest);

  if (prefixedMatch) {
    return prefixedMatch[1];
  }

  return extractSha256FromText(normalizedDigest);
}

export function getDefaultInstallDirectory(
  env: NodeJS.ProcessEnv = process.env,
  platform: NodeJS.Platform = process.platform,
): string {
  if (platform !== "win32") {
    return path.join(os.homedir(), ".local", "share", APP_NAME);
  }

  const localAppData = env.LOCALAPPDATA ?? path.join(os.homedir(), "AppData", "Local");

  return path.join(localAppData, "Programs", APP_NAME);
}

export function getInstalledExecutablePath(installDirectory: string): string {
  return path.join(installDirectory, APP_EXECUTABLE_NAME);
}

export function getStartMenuShortcutPath(
  env: NodeJS.ProcessEnv = process.env,
  platform: NodeJS.Platform = process.platform,
): string {
  const programsDirectory =
    platform === "win32" && env.APPDATA
      ? path.join(env.APPDATA, "Microsoft", "Windows", "Start Menu", "Programs")
      : path.join(os.homedir(), "Start Menu", "Programs");

  return path.join(programsDirectory, `${APP_NAME}.lnk`);
}

export function buildStartMenuShortcutDetails(executablePath: string): InstallerShortcutDetails {
  return {
    cwd: path.dirname(executablePath),
    description: `Start ${APP_NAME}`,
    icon: executablePath,
    iconIndex: 0,
    target: executablePath,
  };
}

export function buildFileOpenCommand(executablePath: string): string {
  return `"${executablePath}" "%1"`;
}

export function buildFileAssociationRegistryEntries(executablePath: string) {
  return [
    {
      key: FILE_ASSOCIATION_EXTENSION_KEY,
      value: FILE_ASSOCIATION_PROG_ID,
    },
    {
      key: FILE_ASSOCIATION_PROG_ID_KEY,
      value: FILE_ASSOCIATION_DESCRIPTION,
    },
    {
      key: `${FILE_ASSOCIATION_PROG_ID_KEY}\\DefaultIcon`,
      value: `"${executablePath}",0`,
    },
    {
      key: `${FILE_ASSOCIATION_PROG_ID_KEY}\\shell\\open\\command`,
      value: buildFileOpenCommand(executablePath),
    },
  ];
}

export function normalizeInstallerOptions(options: Partial<InstallerOptions>): InstallerOptions {
  return {
    fileAssociation: options.fileAssociation === true,
    startMenuShortcut: options.startMenuShortcut === true,
  };
}

export function parseRegistryDefaultString(output: string): string | null {
  for (const line of output.split(/\r?\n/u)) {
    const match = /\sREG_SZ\s+(.*)$/u.exec(line);

    if (match) {
      return match[1].trim();
    }
  }

  return null;
}

export function selectWindowsReleaseAsset(
  release: Pick<LatestReleaseInfo, "assets">,
  arch: NodeJS.Architecture,
): ReleaseAsset | null {
  const archToken = arch === "x64" ? "x64" : arch === "arm64" ? "arm64" : arch;
  const candidates = release.assets.filter((asset) => {
    const name = asset.name.toLowerCase();

    return (
      name.startsWith("spready-") &&
      name.includes("windows") &&
      name.includes(archToken.toLowerCase()) &&
      name.endsWith(".zip")
    );
  });

  return candidates[0] ?? null;
}

export function selectSha256Asset(
  release: Pick<LatestReleaseInfo, "assets">,
  assetName: string,
): ReleaseAsset | null {
  const normalizedAssetName = assetName.toLowerCase();

  return (
    release.assets.find((asset) => asset.name.toLowerCase() === `${normalizedAssetName}.sha256`) ??
    release.assets.find((asset) => {
      const name = asset.name.toLowerCase();

      return name.endsWith(".sha256") && name.includes(normalizedAssetName.replace(/\.zip$/u, ""));
    }) ??
    null
  );
}

export function parseLatestReleaseResponse(response: string): LatestReleaseInfo {
  const parsed = JSON.parse(response) as {
    assets?: Array<Partial<ReleaseAsset>>;
    html_url?: unknown;
    tag_name?: unknown;
  };

  if (typeof parsed.tag_name !== "string") {
    throw new Error("Latest release response did not include a tag_name.");
  }

  const version = parseVersionTag(parsed.tag_name);

  if (!version) {
    throw new Error(`Latest release tag "${parsed.tag_name}" is not a semantic version.`);
  }

  const assets = Array.isArray(parsed.assets)
    ? parsed.assets
        .filter(
          (asset): asset is ReleaseAsset =>
            typeof asset.name === "string" && typeof asset.browser_download_url === "string",
        )
        .map((asset) => ({
          browser_download_url: asset.browser_download_url,
          ...(typeof asset.digest === "string" ? { digest: asset.digest } : {}),
          name: asset.name,
        }))
    : [];

  return {
    assets,
    html_url: typeof parsed.html_url === "string" ? parsed.html_url : "",
    tag_name: parsed.tag_name,
    version,
  };
}

export class InstallerService {
  #arch: NodeJS.Architecture;
  #commandRunner: InstallerCommandRunner;
  #currentAppDirectory: string;
  #currentExecutablePath: string;
  #currentVersion: string;
  #detachedPowerShellRunner: InstallerDetachedPowerShellRunner;
  #env: NodeJS.ProcessEnv;
  #fetch: typeof fetch;
  #isPackaged: boolean;
  #platform: NodeJS.Platform;
  #requestQuit?: () => void;
  #writeShortcut?: (
    shortcutPath: string,
    operation: InstallerShortcutOperation,
    shortcut: InstallerShortcutDetails,
  ) => boolean;

  constructor(dependencies: InstallerServiceDependencies) {
    this.#arch = dependencies.arch ?? process.arch;
    this.#commandRunner = dependencies.commandRunner ?? runProcess;
    this.#currentAppDirectory = dependencies.currentAppDirectory;
    this.#currentExecutablePath = dependencies.currentExecutablePath;
    this.#currentVersion = dependencies.currentVersion;
    this.#detachedPowerShellRunner = dependencies.detachedPowerShellRunner ?? spawnPowerShellScript;
    this.#env = dependencies.env ?? process.env;
    this.#fetch = dependencies.fetch ?? fetch;
    this.#isPackaged = dependencies.isPackaged;
    this.#platform = dependencies.platform ?? process.platform;
    this.#requestQuit = dependencies.requestQuit;
    this.#writeShortcut = dependencies.writeShortcut;
  }

  async getStatus(): Promise<InstallerStatus> {
    const installDirectory = getDefaultInstallDirectory(this.#env, this.#platform);
    const installedExecutablePath = getInstalledExecutablePath(installDirectory);
    const installed = await isRegularFile(installedExecutablePath);
    const canManageInstalledInstance =
      installed &&
      (await pathsReferToSameFile(this.#currentExecutablePath, installedExecutablePath));
    const shortcutPath = getStartMenuShortcutPath(this.#env, this.#platform);

    return {
      canManageInstalledInstance,
      currentVersion: this.#currentVersion,
      installDirectory,
      installed,
      installedExecutablePath,
      isPackaged: this.#isPackaged,
      options: {
        fileAssociation: installed
          ? await this.#isFileAssociationRegistered(installedExecutablePath)
          : false,
        startMenuShortcut: installed ? await isRegularFile(shortcutPath) : false,
      },
      platform: this.#platform,
      updateSupported: this.#platform === "win32",
    };
  }

  async installCurrentApp(options: InstallerOptions): Promise<InstallerOperationResult> {
    this.#assertWindowsPackaged("Installation");

    const status = await this.getStatus();

    await fs.mkdir(status.installDirectory, { recursive: true });

    if (!status.canManageInstalledInstance) {
      const temporaryInstallDirectory = `${status.installDirectory}.tmp-${process.pid}`;

      await fs.rm(temporaryInstallDirectory, { force: true, recursive: true });
      await copyDirectoryContents(this.#currentAppDirectory, temporaryInstallDirectory);
      await fs.rm(status.installDirectory, { force: true, recursive: true });
      await fs.rename(temporaryInstallDirectory, status.installDirectory);
    }

    await this.#applyOptions(options);
    await this.#registerUninstallEntry();

    return {
      message: `${APP_NAME} was installed.`,
      status: await this.getStatus(),
    };
  }

  async applyOptions(options: InstallerOptions): Promise<InstallerOperationResult> {
    const status = await this.getStatus();

    if (!status.installed) {
      throw new Error(`${APP_NAME} is not installed.`);
    }

    await this.#applyOptions(options);

    return {
      message: "Installation options were updated.",
      status: await this.getStatus(),
    };
  }

  async startUninstall(): Promise<InstallerOperationResult> {
    const status = await this.getStatus();

    if (!status.installed) {
      throw new Error(`${APP_NAME} is not installed.`);
    }

    await this.#applyOptions({ fileAssociation: false, startMenuShortcut: false });
    await this.#commandRunner("reg.exe", ["delete", UNINSTALL_REGISTRY_KEY, "/f"]).catch(
      () => undefined,
    );
    const logPath = createInstallerLogPath("spready-uninstall");

    await this.#detachedPowerShellRunner(
      "Uninstall Spready",
      buildUninstallScript(process.pid, status.installDirectory),
      {
        logPath,
      },
    );
    this.#requestQuit?.();

    return {
      logPath,
      message: `${APP_NAME} uninstall started.`,
      status: await this.getStatus(),
    };
  }

  async checkForUpdates(
    request: InstallerCheckUpdatesRequest = {},
  ): Promise<InstallerCheckUpdatesResult> {
    const status = await this.getStatus();

    if (!status.canManageInstalledInstance) {
      return {
        currentVersion: this.#currentVersion,
        message: "Updates can only be checked from the installed Spready executable.",
        status,
        updateAvailable: false,
        updateStarted: false,
      };
    }

    const release = await this.#fetchLatestRelease();
    const currentVersion = parseVersionTag(this.#currentVersion);

    if (!currentVersion) {
      throw new Error(`Current version "${this.#currentVersion}" is not a semantic version.`);
    }

    if (!isVersionNewer(release.version, currentVersion)) {
      return {
        currentVersion: this.#currentVersion,
        latestVersion: formatVersionNumber(release.version),
        message: `${APP_NAME} is up to date.`,
        releaseUrl: release.html_url,
        status,
        updateAvailable: false,
        updateStarted: false,
      };
    }

    const asset = selectWindowsReleaseAsset(release, this.#arch);

    if (!asset) {
      throw new Error(
        "The latest release does not include a Windows bundle for this architecture.",
      );
    }

    if (!request.startUpdate) {
      return {
        assetName: asset.name,
        currentVersion: this.#currentVersion,
        latestVersion: formatVersionNumber(release.version),
        message: `${APP_NAME} ${formatVersionNumber(release.version)} is available.`,
        releaseUrl: release.html_url,
        status,
        updateAvailable: true,
        updateStarted: false,
      };
    }

    const updateDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-update-"));
    const downloadedArchive = path.join(updateDirectory, asset.name);
    const extractionDirectory = path.join(updateDirectory, "extracted");
    const stagingDirectory = path.join(updateDirectory, "stage");
    const expectedSha256 = await this.#resolveAssetSha256(release, asset);

    await this.#downloadFile(
      asset.browser_download_url,
      downloadedArchive,
      MAX_ASSET_DOWNLOAD_BYTES,
    );
    await verifyFileSha256(downloadedArchive, expectedSha256);
    await fs.mkdir(extractionDirectory, { recursive: true });
    const logPath = createInstallerLogPath("update");

    await runPowerShellScript(
      "Extract Spready update",
      `Expand-Archive -LiteralPath ${quotePowerShellString(
        downloadedArchive,
      )} -DestinationPath ${quotePowerShellString(extractionDirectory)} -Force`,
      { logPath },
    );

    const extractedAppDirectory = await findExtractedAppDirectory(extractionDirectory);

    await fs.rm(stagingDirectory, { force: true, recursive: true });
    await fs.rename(extractedAppDirectory, stagingDirectory);

    await this.#detachedPowerShellRunner(
      "Finish Spready update",
      buildFinishUpdateScript({
        installDirectory: status.installDirectory,
        latestVersion: formatVersionNumber(release.version),
        pid: process.pid,
        restart: request.restart ?? true,
        stagedDirectory: stagingDirectory,
        updateDirectory,
      }),
      { logPath },
    );
    this.#requestQuit?.();

    return {
      assetName: asset.name,
      currentVersion: this.#currentVersion,
      latestVersion: formatVersionNumber(release.version),
      logPath,
      message: `${APP_NAME} update started.`,
      releaseUrl: release.html_url,
      status,
      updateAvailable: true,
      updateStarted: true,
    };
  }

  async #applyOptions(options: InstallerOptions) {
    const normalizedOptions = normalizeInstallerOptions(options);
    const installDirectory = getDefaultInstallDirectory(this.#env, this.#platform);
    const shortcutPath = getStartMenuShortcutPath(this.#env, this.#platform);

    if (normalizedOptions.startMenuShortcut) {
      await fs.mkdir(path.dirname(shortcutPath), { recursive: true });

      const installedExecutablePath = getInstalledExecutablePath(installDirectory);
      const shortcut = buildStartMenuShortcutDetails(installedExecutablePath);

      if (!(this.#writeShortcut?.(shortcutPath, "create", shortcut) ?? false)) {
        throw new Error("Failed to write the Start Menu shortcut.");
      }
    } else {
      await fs.rm(shortcutPath, { force: true });
    }

    const installedExecutablePath = getInstalledExecutablePath(installDirectory);

    if (normalizedOptions.fileAssociation) {
      await this.#registerFileAssociation(installedExecutablePath);
    } else {
      await this.#unregisterFileAssociation();
    }
  }

  #assertWindowsPackaged(operation: string) {
    if (this.#platform !== "win32") {
      throw new Error(`${operation} is only supported on Windows.`);
    }

    if (!this.#isPackaged) {
      throw new Error(`${operation} is only available from a packaged Spready build.`);
    }
  }

  async #downloadFile(url: string, destination: string, maxBytes: number) {
    const response = await this.#fetch(url, {
      headers: {
        "User-Agent": "Spready updater",
      },
    });

    if (!response.ok) {
      throw new Error(`Download failed with HTTP ${response.status}.`);
    }

    const buffer = Buffer.from(await response.arrayBuffer());

    if (buffer.byteLength > maxBytes) {
      throw new Error("Downloaded update asset is larger than expected.");
    }

    await fs.writeFile(destination, buffer);
  }

  async #fetchLatestRelease() {
    const response = await this.#fetch(GITHUB_LATEST_RELEASE_URL, {
      headers: {
        Accept: "application/vnd.github+json",
        "User-Agent": "Spready updater",
        "X-GitHub-Api-Version": "2022-11-28",
      },
    });

    if (!response.ok) {
      throw new Error(`GitHub latest release request failed with HTTP ${response.status}.`);
    }

    const text = await response.text();

    if (Buffer.byteLength(text, "utf8") > MAX_RELEASE_RESPONSE_BYTES) {
      throw new Error("GitHub latest release response is larger than expected.");
    }

    return parseLatestReleaseResponse(text);
  }

  async #registerUninstallEntry() {
    const status = await this.getStatus();
    const uninstallCommand = `"${status.installedExecutablePath}" --spready-uninstall`;

    await this.#commandRunner("reg.exe", ["add", UNINSTALL_REGISTRY_KEY, "/f"]);
    await this.#setRegistryString(UNINSTALL_REGISTRY_KEY, "DisplayName", APP_NAME);
    await this.#setRegistryString(UNINSTALL_REGISTRY_KEY, "DisplayVersion", this.#currentVersion);
    await this.#setRegistryString(UNINSTALL_REGISTRY_KEY, "Publisher", PUBLISHER);
    await this.#setRegistryString(
      UNINSTALL_REGISTRY_KEY,
      "InstallLocation",
      status.installDirectory,
    );
    await this.#setRegistryString(
      UNINSTALL_REGISTRY_KEY,
      "DisplayIcon",
      status.installedExecutablePath,
    );
    await this.#setRegistryString(UNINSTALL_REGISTRY_KEY, "UninstallString", uninstallCommand);
    await this.#setRegistryString(
      UNINSTALL_REGISTRY_KEY,
      "QuietUninstallString",
      `${uninstallCommand} --quiet`,
    );
    await this.#setRegistryDword(UNINSTALL_REGISTRY_KEY, "NoModify", 1);
    await this.#setRegistryDword(UNINSTALL_REGISTRY_KEY, "NoRepair", 1);
  }

  async #resolveAssetSha256(release: LatestReleaseInfo, asset: ReleaseAsset) {
    const digestSha256 = extractSha256FromDigest(asset.digest);

    if (digestSha256) {
      return digestSha256;
    }

    const shaAsset = selectSha256Asset(release, asset.name);

    if (!shaAsset) {
      throw new Error("The update asset did not include a SHA-256 digest.");
    }

    const response = await this.#fetch(shaAsset.browser_download_url, {
      headers: {
        "User-Agent": "Spready updater",
      },
    });

    if (!response.ok) {
      throw new Error(`SHA-256 download failed with HTTP ${response.status}.`);
    }

    const text = await response.text();
    const sha256 = extractSha256FromText(text);

    if (!sha256) {
      throw new Error("Downloaded SHA-256 file did not include a valid digest.");
    }

    return sha256;
  }

  async #isFileAssociationRegistered(executablePath: string) {
    if (this.#platform !== "win32") {
      return false;
    }

    const extensionProgId = await this.#queryRegistryDefaultValue(FILE_ASSOCIATION_EXTENSION_KEY);

    if (extensionProgId !== FILE_ASSOCIATION_PROG_ID) {
      return false;
    }

    const openCommand = await this.#queryRegistryDefaultValue(
      `${FILE_ASSOCIATION_PROG_ID_KEY}\\shell\\open\\command`,
    );

    return openCommand === buildFileOpenCommand(executablePath);
  }

  async #queryRegistryDefaultValue(key: string) {
    try {
      const result = await this.#commandRunner("reg.exe", ["query", key, "/ve"]);

      return parseRegistryDefaultString(result.stdout);
    } catch {
      return null;
    }
  }

  async #registerFileAssociation(executablePath: string) {
    for (const entry of buildFileAssociationRegistryEntries(executablePath)) {
      await this.#setRegistryDefaultString(entry.key, entry.value);
    }
  }

  async #unregisterFileAssociation() {
    const extensionProgId = await this.#queryRegistryDefaultValue(FILE_ASSOCIATION_EXTENSION_KEY);

    if (extensionProgId === FILE_ASSOCIATION_PROG_ID) {
      await this.#commandRunner("reg.exe", [
        "delete",
        FILE_ASSOCIATION_EXTENSION_KEY,
        "/ve",
        "/f",
      ]).catch(() => undefined);
    }

    await this.#commandRunner("reg.exe", ["delete", FILE_ASSOCIATION_PROG_ID_KEY, "/f"]).catch(
      () => undefined,
    );
  }

  async #setRegistryDefaultString(key: string, value: string) {
    await this.#commandRunner("reg.exe", ["add", key, "/ve", "/t", "REG_SZ", "/d", value, "/f"]);
  }

  async #setRegistryString(key: string, name: string, value: string) {
    await this.#commandRunner("reg.exe", [
      "add",
      key,
      "/v",
      name,
      "/t",
      "REG_SZ",
      "/d",
      value,
      "/f",
    ]);
  }

  async #setRegistryDword(key: string, name: string, value: number) {
    await this.#commandRunner("reg.exe", [
      "add",
      key,
      "/v",
      name,
      "/t",
      "REG_DWORD",
      "/d",
      String(value),
      "/f",
    ]);
  }
}

export async function runWithAsarFilesystemDisabled<T>(operation: () => Promise<T>): Promise<T> {
  const processWithAsarFlag = process as NodeJS.Process & {
    noAsar?: boolean;
  };
  const previousNoAsar = processWithAsarFlag.noAsar;

  processWithAsarFlag.noAsar = true;

  try {
    return await operation();
  } finally {
    if (previousNoAsar === undefined) {
      Reflect.deleteProperty(processWithAsarFlag, "noAsar");
    } else {
      processWithAsarFlag.noAsar = previousNoAsar;
    }
  }
}

async function copyDirectoryContents(sourceDirectory: string, targetDirectory: string) {
  await runWithAsarFilesystemDisabled(async () => {
    await fs.mkdir(targetDirectory, { recursive: true });

    for (const entry of await fs.readdir(sourceDirectory)) {
      await fs.cp(path.join(sourceDirectory, entry), path.join(targetDirectory, entry), {
        dereference: false,
        recursive: true,
      });
    }
  });
}

async function findExtractedAppDirectory(extractionDirectory: string) {
  if (await isRegularFile(path.join(extractionDirectory, APP_EXECUTABLE_NAME))) {
    return extractionDirectory;
  }

  const entries = await fs.readdir(extractionDirectory, { withFileTypes: true });

  for (const entry of entries) {
    if (!entry.isDirectory()) {
      continue;
    }

    const candidate = path.join(extractionDirectory, entry.name);

    if (await isRegularFile(path.join(candidate, APP_EXECUTABLE_NAME))) {
      return candidate;
    }
  }

  throw new Error("The downloaded update archive did not contain Spready.exe.");
}

async function isRegularFile(filePath: string): Promise<boolean> {
  try {
    return (await fs.stat(filePath)).isFile();
  } catch {
    return false;
  }
}

async function pathsReferToSameFile(left: string, right: string): Promise<boolean> {
  try {
    const [leftRealPath, rightRealPath] = await Promise.all([
      fs.realpath(left),
      fs.realpath(right),
    ]);

    return leftRealPath.toLowerCase() === rightRealPath.toLowerCase();
  } catch {
    return path.resolve(left).toLowerCase() === path.resolve(right).toLowerCase();
  }
}

async function runProcess(command: string, args: string[]): Promise<InstallerCommandResult> {
  return await new Promise<InstallerCommandResult>((resolve, reject) => {
    const child = spawn(command, args, {
      windowsHide: true,
    });
    let stderr = "";
    let stdout = "";

    child.stdout?.on("data", (chunk: Buffer) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr?.on("data", (chunk: Buffer) => {
      stderr += chunk.toString("utf8");
    });

    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve({ stderr, stdout });
        return;
      }

      reject(new Error(stderr.trim() || `${command} exited with code ${code ?? "unknown"}.`));
    });
  });
}

async function runPowerShellScript(
  operationName: string,
  script: string,
  options: { logPath?: string } = {},
) {
  const logPath = options.logPath ?? createInstallerLogPath(operationName);
  const scriptPath = createInstallerScriptPath(logPath);
  const wrappedScript = buildInstallerPowerShellScript(operationName, script, {
    logPath,
    terminalBehavior: "exit",
  });

  try {
    await writePowerShellScriptFile(scriptPath, wrappedScript);
    await runProcess("powershell.exe", buildPowerShellFileArguments(scriptPath));
    await fs.rm(scriptPath, { force: true });
  } catch (error) {
    const message = error instanceof Error ? error.message : "PowerShell script failed.";

    throw new Error(`${message} Log file: ${logPath} Script file: ${scriptPath}`);
  }
}

async function spawnPowerShellScript(
  operationName: string,
  script: string,
  options: { logPath?: string } = {},
) {
  const logPath = options.logPath ?? createInstallerLogPath(operationName);
  const powerShellLogPath = getInstallerPowerShellLogPath(logPath);
  const scriptPath = createInstallerScriptPath(logPath);
  const wrappedScript = buildInstallerPowerShellScript(operationName, script, {
    logPath: powerShellLogPath,
    terminalBehavior: "exit",
  });

  await writePowerShellScriptFile(scriptPath, wrappedScript);
  await appendInstallerLogLine(logPath, "INFO", `Launching detached PowerShell: ${operationName}`);
  await appendInstallerLogLine(logPath, "INFO", `PowerShell script file: ${scriptPath}`);
  await appendInstallerLogLine(logPath, "INFO", `PowerShell log file: ${powerShellLogPath}`);

  let child: ReturnType<typeof spawn>;

  try {
    child = spawn("cmd.exe", buildDetachedPowerShellStartArguments(scriptPath), {
      stdio: "ignore",
      windowsHide: true,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown spawn error.";

    await appendInstallerLogLine(logPath, "ERROR", `Failed to launch PowerShell: ${message}`);
    throw error;
  }

  child.on("exit", (code) => {
    if (code === 0) {
      return;
    }

    void appendInstallerLogLine(
      logPath,
      "ERROR",
      `PowerShell launcher exited with code ${code ?? "unknown"}.`,
    );
  });
  const handoff = waitForDetachedProcessHandoff(child, logPath, powerShellLogPath);

  try {
    await handoff;
  } finally {
    child.unref();
  }
}

async function waitForDetachedProcessHandoff(
  child: ReturnType<typeof spawn>,
  logPath: string,
  powerShellLogPath: string,
): Promise<void> {
  await new Promise<void>((resolve) => {
    const timeout = setTimeout(resolve, 500);
    const finish = () => {
      clearTimeout(timeout);
      resolve();
    };

    child.once("spawn", () => {
      void appendInstallerLogLine(logPath, "INFO", "Detached PowerShell process started.").finally(
        finish,
      );
    });
    child.once("error", (error) => {
      void appendInstallerLogLine(
        logPath,
        "ERROR",
        `Failed to launch PowerShell terminal: ${error.message}`,
      ).finally(finish);
    });
  });

  await waitForInstallerScriptStartup(logPath, powerShellLogPath);
}

async function waitForInstallerScriptStartup(
  logPath: string,
  powerShellLogPath: string,
  timeoutMilliseconds = 5000,
): Promise<void> {
  const deadline = Date.now() + timeoutMilliseconds;

  while (Date.now() < deadline) {
    if (await isRegularFile(powerShellLogPath)) {
      return;
    }

    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  const message = `Detached PowerShell script did not start within ${timeoutMilliseconds} ms.`;

  await appendInstallerLogLine(logPath, "ERROR", message);
  throw new Error(`${message} Log file: ${logPath} PowerShell log file: ${powerShellLogPath}`);
}

function createInstallerLogPath(operationName: string) {
  const safeOperationName = operationName
    .toLowerCase()
    .replace(/[^a-z0-9]+/gu, "-")
    .replace(/^-|-$/gu, "");
  const timestamp = new Date().toISOString().replace(/[:.]/gu, "-");

  return path.join(
    os.tmpdir(),
    INSTALLER_LOG_DIRECTORY_NAME,
    `${timestamp}-${process.pid}-${safeOperationName || "operation"}.log`,
  );
}

function createInstallerScriptPath(logPath: string) {
  return logPath.toLowerCase().endsWith(".log") ? `${logPath.slice(0, -4)}.ps1` : `${logPath}.ps1`;
}

export function getInstallerPowerShellLogPath(logPath: string) {
  return logPath.toLowerCase().endsWith(".log")
    ? `${logPath.slice(0, -4)}.powershell.log`
    : `${logPath}.powershell.log`;
}

export function buildPowerShellFileArguments(
  scriptPath: string,
  options: { keepOpen?: boolean } = {},
) {
  return [
    ...(options.keepOpen ? ["-NoExit"] : []),
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
  ];
}

export function buildDetachedPowerShellStartArguments(scriptPath: string) {
  return [
    "/d",
    "/s",
    "/c",
    "start",
    '""',
    "/min",
    "/D",
    path.dirname(scriptPath),
    "powershell.exe",
    ...buildPowerShellFileArguments(scriptPath),
  ];
}

async function writePowerShellScriptFile(scriptPath: string, script: string) {
  await fs.mkdir(path.dirname(scriptPath), { recursive: true });
  await fs.writeFile(
    scriptPath,
    Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(script, "utf16le")]),
  );
}

async function appendInstallerLogLine(logPath: string, level: "ERROR" | "INFO", message: string) {
  const timestamp = new Date().toISOString();

  try {
    await fs.mkdir(path.dirname(logPath), { recursive: true });
    await fs.appendFile(logPath, `[${timestamp}] [${level}] ${message}\n`);
  } catch {
    // Launch failure logging should never take down the app.
  }
}

export function buildInstallerPowerShellScript(
  operationName: string,
  script: string,
  options: { logPath: string; terminalBehavior: "exit" | "keep-open" },
) {
  return [
    "$ErrorActionPreference = 'Stop'",
    "$ProgressPreference = 'Continue'",
    `$spreadyInstallerLogPath = ${quotePowerShellString(options.logPath)}`,
    `$spreadyInstallerOperation = ${quotePowerShellString(operationName)}`,
    "New-Item -ItemType Directory -Path (Split-Path -Parent $spreadyInstallerLogPath) -Force | Out-Null",
    "function Write-SpreadyInstallerLog {",
    "  param([string] $Message, [string] $Level = 'INFO')",
    "  $timestamp = Get-Date -Format o",
    '  $line = "[$timestamp] [$Level] $Message"',
    "  Write-Host $line",
    "  for ($attempt = 1; $attempt -le 20; $attempt += 1) {",
    "    try {",
    "      Add-Content -LiteralPath $spreadyInstallerLogPath -Value $line",
    "      return",
    "    } catch {",
    "      if ($attempt -eq 20) {",
    '        Write-Host "[$timestamp] [ERROR] Failed to write installer log after ${attempt} attempts: $($_.Exception.Message)"',
    "        return",
    "      }",
    "      Start-Sleep -Milliseconds 50",
    "    }",
    "  }",
    "}",
    "function Write-SpreadyInstallerError {",
    "  param([System.Management.Automation.ErrorRecord] $ErrorRecord)",
    '  Write-SpreadyInstallerLog "Error: $($ErrorRecord.Exception.Message)" "ERROR"',
    "  if ($ErrorRecord.InvocationInfo.PositionMessage) {",
    '    Write-SpreadyInstallerLog $ErrorRecord.InvocationInfo.PositionMessage "ERROR"',
    "  }",
    "  if ($ErrorRecord.ScriptStackTrace) {",
    '    Write-SpreadyInstallerLog $ErrorRecord.ScriptStackTrace "ERROR"',
    "  }",
    "}",
    "function Invoke-SpreadyInstallerStep {",
    "  param([string] $Name, [scriptblock] $Action)",
    '  Write-SpreadyInstallerLog "Starting: $Name"',
    "  try {",
    "    & $Action",
    '    Write-SpreadyInstallerLog "Completed: $Name"',
    "  } catch {",
    '    Write-SpreadyInstallerLog "Failed: $Name" "ERROR"',
    "    Write-SpreadyInstallerError $_",
    "    throw",
    "  }",
    "}",
    'Write-SpreadyInstallerLog "Starting $spreadyInstallerOperation."',
    'Write-SpreadyInstallerLog "Log file: $spreadyInstallerLogPath"',
    "$spreadyInstallerExitCode = 0",
    "try {",
    indentPowerShellScript(script),
    '  Write-SpreadyInstallerLog "Finished $spreadyInstallerOperation."',
    "} catch {",
    "  $spreadyInstallerExitCode = 1",
    '  Write-SpreadyInstallerLog "$spreadyInstallerOperation failed." "ERROR"',
    "  Write-SpreadyInstallerError $_",
    "}",
    options.terminalBehavior === "exit"
      ? "exit $spreadyInstallerExitCode"
      : 'Write-SpreadyInstallerLog "PowerShell window left open for review. Close it when finished."',
  ].join("\n");
}

function quotePowerShellString(value: string) {
  return `'${value.replaceAll("'", "''")}'`;
}

export function buildWaitForProcessExitScript(pid: number, timeoutSeconds = 60) {
  return [
    `$deadline = (Get-Date).AddSeconds(${timeoutSeconds})`,
    `while ((Get-Process -Id ${pid} -ErrorAction SilentlyContinue) -and ((Get-Date) -lt $deadline)) {`,
    "  Start-Sleep -Milliseconds 250",
    "}",
    `if (Get-Process -Id ${pid} -ErrorAction SilentlyContinue) {`,
    `  throw '${APP_NAME} did not exit within ${timeoutSeconds} seconds.'`,
    "}",
  ].join("\n");
}

export function buildWaitForInstallProcessesExitScript(
  installDirectory: string,
  timeoutSeconds = 120,
) {
  return [
    `$installRoot = [System.IO.Path]::GetFullPath(${quotePowerShellString(
      installDirectory,
    )}).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar`,
    `$deadline = (Get-Date).AddSeconds(${timeoutSeconds})`,
    "$lastProcessSummary = ''",
    "while ((Get-Date) -lt $deadline) {",
    "  $runningProcesses = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {",
    "    $_.ExecutablePath -and $_.ExecutablePath.StartsWith($installRoot, [System.StringComparison]::OrdinalIgnoreCase)",
    "  })",
    "  if ($runningProcesses.Count -eq 0) {",
    "    return",
    "  }",
    "  $lastProcessSummary = (($runningProcesses | ForEach-Object { \"$($_.Name)[$($_.ProcessId)]\" }) -join ', ')",
    "  Start-Sleep -Milliseconds 500",
    "}",
    `throw (${quotePowerShellString(
      `${APP_NAME} processes are still running from the install directory: `,
    )} + $lastProcessSummary)`,
  ].join("\n");
}

export function buildUninstallScript(pid: number, installDirectory: string) {
  return [
    "Invoke-SpreadyInstallerStep 'wait for Spready to exit' {",
    indentPowerShellScript(buildWaitForProcessExitScript(pid)),
    "}",
    "Invoke-SpreadyInstallerStep 'remove install directory' {",
    "  $lastRemoveError = $null",
    "  for ($attempt = 1; $attempt -le 60; $attempt += 1) {",
    `    if (-not (Test-Path -LiteralPath ${quotePowerShellString(installDirectory)})) {`,
    "      if ($attempt -eq 1) {",
    "        Write-SpreadyInstallerLog 'Install directory was already removed.'",
    "      } else {",
    '        Write-SpreadyInstallerLog "Install directory removed after $($attempt - 1) failed attempt(s)."',
    "      }",
    "      return",
    "    }",
    "    try {",
    `      Remove-Item -LiteralPath ${quotePowerShellString(installDirectory)} -Recurse -Force`,
    '      Write-SpreadyInstallerLog "Install directory removed on attempt $attempt."',
    "      return",
    "    } catch {",
    "      $lastRemoveError = $_",
    '      Write-SpreadyInstallerLog "Install directory removal failed on attempt ${attempt}: $($_.Exception.Message)" "ERROR"',
    "      Start-Sleep -Milliseconds 500",
    "    }",
    "  }",
    "  Write-SpreadyInstallerError $lastRemoveError",
    "  throw $lastRemoveError",
    "}",
  ].join("\n");
}

function indentPowerShellScript(script: string) {
  return script
    .split("\n")
    .map((line) => `  ${line}`)
    .join("\n");
}

export function buildFinishUpdateScript(args: {
  installDirectory: string;
  latestVersion: string;
  pid: number;
  restart: boolean;
  stagedDirectory: string;
  updateDirectory: string;
}) {
  const installedExecutablePath = getInstalledExecutablePath(args.installDirectory);
  const backupDirectory = `${args.installDirectory}.old-${Date.now()}`;

  return [
    "function Invoke-SpreadyUpdateStep {",
    "  param([string] $Name, [scriptblock] $Action)",
    "  $lastError = $null",
    '  Write-SpreadyInstallerLog "Starting: $Name"',
    "  for ($attempt = 1; $attempt -le 40; $attempt += 1) {",
    "    try {",
    "      & $Action",
    '      Write-SpreadyInstallerLog "Completed: $Name"',
    "      return",
    "    } catch {",
    "      $lastError = $_",
    '      Write-SpreadyInstallerLog "$Name failed on attempt ${attempt}: $($_.Exception.Message)" "ERROR"',
    "      Start-Sleep -Milliseconds 500",
    "    }",
    "  }",
    '  Write-SpreadyInstallerLog "Failed: $Name" "ERROR"',
    "  Write-SpreadyInstallerError $lastError",
    "  throw $lastError",
    "}",
    "try {",
    "  Invoke-SpreadyInstallerStep 'wait for Spready to exit' {",
    indentPowerShellScript(buildWaitForProcessExitScript(args.pid)),
    "  }",
    "  Invoke-SpreadyInstallerStep 'wait for installed Spready processes to exit' {",
    indentPowerShellScript(buildWaitForInstallProcessesExitScript(args.installDirectory)),
    "  }",
    `  if (Test-Path -LiteralPath ${quotePowerShellString(backupDirectory)}) {`,
    `    Invoke-SpreadyUpdateStep 'remove stale backup' { Remove-Item -LiteralPath ${quotePowerShellString(
      backupDirectory,
    )} -Recurse -Force }`,
    "  }",
    `  if (Test-Path -LiteralPath ${quotePowerShellString(args.installDirectory)}) {`,
    `    Invoke-SpreadyUpdateStep 'rename current install' { Rename-Item -LiteralPath ${quotePowerShellString(
      args.installDirectory,
    )} -NewName ${quotePowerShellString(path.basename(backupDirectory))} -Force }`,
    "  }",
    `  Invoke-SpreadyUpdateStep 'move staged update' { Move-Item -LiteralPath ${quotePowerShellString(
      args.stagedDirectory,
    )} -Destination ${quotePowerShellString(args.installDirectory)} -Force }`,
    `  if (Test-Path -LiteralPath ${quotePowerShellString(backupDirectory)}) {`,
    "    try {",
    `      Invoke-SpreadyUpdateStep 'remove old install backup' { Remove-Item -LiteralPath ${quotePowerShellString(
      backupDirectory,
    )} -Recurse -Force }`,
    "    } catch {",
    '      Write-SpreadyInstallerLog "Old install cleanup failed: $($_.Exception.Message)" "ERROR"',
    "    }",
    "  }",
    "  try {",
    `    & reg.exe add ${quotePowerShellString(UNINSTALL_REGISTRY_KEY)} /v DisplayVersion /t REG_SZ /d ${quotePowerShellString(
      args.latestVersion,
    )} /f | Out-Null`,
    "  } catch {",
    '    Write-SpreadyInstallerLog "Registry version update failed: $($_.Exception.Message)" "ERROR"',
    "  }",
    args.restart
      ? `  Invoke-SpreadyInstallerStep 'restart Spready' { Start-Process -FilePath ${quotePowerShellString(
          installedExecutablePath,
        )} }`
      : "",
    "  try {",
    `    Invoke-SpreadyUpdateStep 'remove temporary update files' { Remove-Item -LiteralPath ${quotePowerShellString(
      args.updateDirectory,
    )} -Recurse -Force }`,
    "  } catch {",
    '    Write-SpreadyInstallerLog "Temporary update cleanup failed: $($_.Exception.Message)" "ERROR"',
    "  }",
    "} catch {",
    '  Write-SpreadyInstallerLog "Update failed: $($_.Exception.Message)" "ERROR"',
    `  if ((Test-Path -LiteralPath ${quotePowerShellString(
      backupDirectory,
    )}) -and -not (Test-Path -LiteralPath ${quotePowerShellString(args.installDirectory)})) {`,
    "    try {",
    `      Invoke-SpreadyUpdateStep 'restore previous install' { Rename-Item -LiteralPath ${quotePowerShellString(
      backupDirectory,
    )} -NewName ${quotePowerShellString(path.basename(args.installDirectory))} -Force }`,
    "      Write-SpreadyInstallerLog 'Restored previous install after update failure.'",
    "    } catch {",
    '      Write-SpreadyInstallerLog "Previous install restore failed: $($_.Exception.Message)" "ERROR"',
    "    }",
    "  }",
    "  throw",
    "}",
  ]
    .filter(Boolean)
    .join("\n");
}

async function verifyFileSha256(filePath: string, expectedSha256: string) {
  const content = await fs.readFile(filePath);
  const actualSha256 = createHash("sha256").update(content).digest("hex");

  if (actualSha256.toLowerCase() !== expectedSha256.toLowerCase()) {
    throw new Error("Downloaded update SHA-256 did not match the expected digest.");
  }
}
