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
const GITHUB_LATEST_RELEASE_URL = "https://api.github.com/repos/tuncb/spready/releases/latest";
const PUBLISHER = "tuncb";
const UNINSTALL_REGISTRY_KEY =
  "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Spready";
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
  currentAppDirectory: string;
  currentExecutablePath: string;
  currentVersion: string;
  env?: NodeJS.ProcessEnv;
  fetch?: typeof fetch;
  getAutoStartEnabled?: (executablePath: string) => boolean;
  isPackaged: boolean;
  platform?: NodeJS.Platform;
  requestQuit?: () => void;
  setAutoStart?: (executablePath: string, enabled: boolean) => void;
  writeShortcut?: (shortcutPath: string, executablePath: string) => boolean;
}

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
  #currentAppDirectory: string;
  #currentExecutablePath: string;
  #currentVersion: string;
  #env: NodeJS.ProcessEnv;
  #fetch: typeof fetch;
  #getAutoStartEnabled?: (executablePath: string) => boolean;
  #isPackaged: boolean;
  #platform: NodeJS.Platform;
  #requestQuit?: () => void;
  #setAutoStart?: (executablePath: string, enabled: boolean) => void;
  #writeShortcut?: (shortcutPath: string, executablePath: string) => boolean;

  constructor(dependencies: InstallerServiceDependencies) {
    this.#arch = dependencies.arch ?? process.arch;
    this.#currentAppDirectory = dependencies.currentAppDirectory;
    this.#currentExecutablePath = dependencies.currentExecutablePath;
    this.#currentVersion = dependencies.currentVersion;
    this.#env = dependencies.env ?? process.env;
    this.#fetch = dependencies.fetch ?? fetch;
    this.#getAutoStartEnabled = dependencies.getAutoStartEnabled;
    this.#isPackaged = dependencies.isPackaged;
    this.#platform = dependencies.platform ?? process.platform;
    this.#requestQuit = dependencies.requestQuit;
    this.#setAutoStart = dependencies.setAutoStart;
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
        autoStart: installed
          ? (this.#getAutoStartEnabled?.(installedExecutablePath) ?? false)
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

    await this.#applyOptions({ autoStart: false, startMenuShortcut: false });
    await runProcess("reg.exe", ["delete", UNINSTALL_REGISTRY_KEY, "/f"]).catch(() => undefined);
    spawnPowerShellScript(buildUninstallScript(process.pid, status.installDirectory));
    this.#requestQuit?.();

    return {
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
    await runPowerShellScript(
      `Expand-Archive -LiteralPath ${quotePowerShellString(
        downloadedArchive,
      )} -DestinationPath ${quotePowerShellString(extractionDirectory)} -Force`,
    );

    const extractedAppDirectory = await findExtractedAppDirectory(extractionDirectory);

    await fs.rm(stagingDirectory, { force: true, recursive: true });
    await fs.rename(extractedAppDirectory, stagingDirectory);

    spawnPowerShellScript(
      buildFinishUpdateScript({
        installDirectory: status.installDirectory,
        pid: process.pid,
        restart: request.restart ?? true,
        stagedDirectory: stagingDirectory,
        updateDirectory,
      }),
    );
    this.#requestQuit?.();

    return {
      assetName: asset.name,
      currentVersion: this.#currentVersion,
      latestVersion: formatVersionNumber(release.version),
      message: `${APP_NAME} update started.`,
      releaseUrl: release.html_url,
      status,
      updateAvailable: true,
      updateStarted: true,
    };
  }

  async #applyOptions(options: InstallerOptions) {
    const installDirectory = getDefaultInstallDirectory(this.#env, this.#platform);
    const installedExecutablePath = getInstalledExecutablePath(installDirectory);
    const shortcutPath = getStartMenuShortcutPath(this.#env, this.#platform);

    this.#setAutoStart?.(installedExecutablePath, options.autoStart);

    if (options.startMenuShortcut) {
      await fs.mkdir(path.dirname(shortcutPath), { recursive: true });

      if (!(this.#writeShortcut?.(shortcutPath, installedExecutablePath) ?? false)) {
        throw new Error("Failed to write the Start Menu shortcut.");
      }
    } else {
      await fs.rm(shortcutPath, { force: true });
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

    await runProcess("reg.exe", ["add", UNINSTALL_REGISTRY_KEY, "/f"]);
    await setRegistryString("DisplayName", APP_NAME);
    await setRegistryString("DisplayVersion", this.#currentVersion);
    await setRegistryString("Publisher", PUBLISHER);
    await setRegistryString("InstallLocation", status.installDirectory);
    await setRegistryString("DisplayIcon", status.installedExecutablePath);
    await setRegistryString("UninstallString", uninstallCommand);
    await setRegistryString("QuietUninstallString", `${uninstallCommand} --quiet`);
    await setRegistryDword("NoModify", 1);
    await setRegistryDword("NoRepair", 1);
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
}

async function copyDirectoryContents(sourceDirectory: string, targetDirectory: string) {
  await fs.mkdir(targetDirectory, { recursive: true });

  for (const entry of await fs.readdir(sourceDirectory)) {
    await fs.cp(path.join(sourceDirectory, entry), path.join(targetDirectory, entry), {
      dereference: false,
      recursive: true,
    });
  }
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

async function runProcess(command: string, args: string[]) {
  await new Promise<void>((resolve, reject) => {
    const child = spawn(command, args, {
      windowsHide: true,
    });
    let stderr = "";

    child.stderr?.on("data", (chunk: Buffer) => {
      stderr += chunk.toString("utf8");
    });

    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }

      reject(new Error(stderr.trim() || `${command} exited with code ${code ?? "unknown"}.`));
    });
  });
}

async function runPowerShellScript(script: string) {
  await runProcess("powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-EncodedCommand",
    encodePowerShellCommand(script),
  ]);
}

function spawnPowerShellScript(script: string) {
  const child = spawn(
    "powershell.exe",
    [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-EncodedCommand",
      encodePowerShellCommand(script),
    ],
    {
      detached: true,
      stdio: "ignore",
      windowsHide: true,
    },
  );

  child.unref();
}

function encodePowerShellCommand(script: string) {
  return Buffer.from(script, "utf16le").toString("base64");
}

function quotePowerShellString(value: string) {
  return `'${value.replaceAll("'", "''")}'`;
}

function buildUninstallScript(pid: number, installDirectory: string) {
  return [
    "$ErrorActionPreference = 'SilentlyContinue'",
    `Wait-Process -Id ${pid} -Timeout 60`,
    `Remove-Item -LiteralPath ${quotePowerShellString(installDirectory)} -Recurse -Force`,
  ].join("\n");
}

function buildFinishUpdateScript(args: {
  installDirectory: string;
  pid: number;
  restart: boolean;
  stagedDirectory: string;
  updateDirectory: string;
}) {
  const installedExecutablePath = getInstalledExecutablePath(args.installDirectory);
  const backupDirectory = `${args.installDirectory}.old-${Date.now()}`;

  return [
    "$ErrorActionPreference = 'Stop'",
    `Wait-Process -Id ${args.pid} -Timeout 60`,
    `if (Test-Path -LiteralPath ${quotePowerShellString(backupDirectory)}) { Remove-Item -LiteralPath ${quotePowerShellString(
      backupDirectory,
    )} -Recurse -Force }`,
    `if (Test-Path -LiteralPath ${quotePowerShellString(args.installDirectory)}) { Rename-Item -LiteralPath ${quotePowerShellString(
      args.installDirectory,
    )} -NewName ${quotePowerShellString(path.basename(backupDirectory))} -Force }`,
    `Move-Item -LiteralPath ${quotePowerShellString(args.stagedDirectory)} -Destination ${quotePowerShellString(
      args.installDirectory,
    )} -Force`,
    `if (Test-Path -LiteralPath ${quotePowerShellString(backupDirectory)}) { Remove-Item -LiteralPath ${quotePowerShellString(
      backupDirectory,
    )} -Recurse -Force }`,
    args.restart ? `Start-Process -FilePath ${quotePowerShellString(installedExecutablePath)}` : "",
    `Remove-Item -LiteralPath ${quotePowerShellString(args.updateDirectory)} -Recurse -Force`,
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

async function setRegistryString(name: string, value: string) {
  await runProcess("reg.exe", [
    "add",
    UNINSTALL_REGISTRY_KEY,
    "/v",
    name,
    "/t",
    "REG_SZ",
    "/d",
    value,
    "/f",
  ]);
}

async function setRegistryDword(name: string, value: number) {
  await runProcess("reg.exe", [
    "add",
    UNINSTALL_REGISTRY_KEY,
    "/v",
    name,
    "/t",
    "REG_DWORD",
    "/d",
    String(value),
    "/f",
  ]);
}
