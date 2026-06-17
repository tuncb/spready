import { readdir, readFile, realpath, stat } from "node:fs/promises";
import path from "node:path";

export interface ManualReadToolHint {
  arguments: {
    path: string;
  };
  name: "read_manual";
}

export interface ManualEntry {
  absolutePath: string;
  path: string;
  readTool: ManualReadToolHint;
  title?: string;
}

export interface ManualListResult {
  manuals: ManualEntry[];
  manualsDirectory: string;
}

export interface ManualReadResult {
  path: string;
  text: string;
  truncated: boolean;
}

export interface ManualDirectoryOptions {
  cwd?: string;
  executablePath?: string;
  manualsDir?: string;
}

export const DEFAULT_MANUAL_READ_MAX_BYTES = 60_000;
export const MAX_MANUAL_READ_BYTES = 200_000;

function toManualPath(filePath: string) {
  return filePath.split(path.sep).join("/");
}

function isAbsolutePath(filePath: string) {
  return (
    path.isAbsolute(filePath) || path.win32.isAbsolute(filePath) || path.posix.isAbsolute(filePath)
  );
}

async function isDirectory(directoryPath: string) {
  try {
    return (await stat(directoryPath)).isDirectory();
  } catch {
    return false;
  }
}

export async function findManualsDirectory(options: ManualDirectoryOptions = {}) {
  const configuredDirectory = options.manualsDir ?? process.env.SPREADY_MCP_MANUALS_DIR;

  if (configuredDirectory) {
    return path.resolve(configuredDirectory);
  }

  const executablePath = options.executablePath ?? process.execPath;
  const cwd = options.cwd ?? process.cwd();
  const candidates = [
    path.join(path.dirname(executablePath), "manuals"),
    path.join(cwd, "manuals"),
  ];

  for (const candidate of candidates) {
    if (await isDirectory(candidate)) {
      return path.resolve(candidate);
    }
  }

  return path.resolve(candidates[0]);
}

export function normalizeManualRequestPath(requestedPath: string) {
  const slashPath = requestedPath.replaceAll("\\", "/").trim();

  if (!slashPath) {
    throw new Error("Manual path is required.");
  }

  if (isAbsolutePath(slashPath)) {
    throw new Error("Manual path must be relative.");
  }

  const normalized = path.posix.normalize(slashPath);

  if (normalized === "." || normalized === ".." || normalized.startsWith("../")) {
    throw new Error("Manual path cannot escape the manuals folder.");
  }

  return normalized;
}

function isWithinRoot(root: string, target: string) {
  const relativePath = path.relative(root, target);

  return relativePath === "" || (!relativePath.startsWith("..") && !isAbsolutePath(relativePath));
}

export async function resolveManualFilePath(manualsDirectory: string, requestedPath: string) {
  const normalizedPath = normalizeManualRequestPath(requestedPath);
  const root = await realpath(manualsDirectory);
  const target = await realpath(path.join(root, normalizedPath));

  if (!isWithinRoot(root, target)) {
    throw new Error("Manual path cannot escape the manuals folder.");
  }

  const fileStat = await stat(target);

  if (!fileStat.isFile()) {
    throw new Error("Manual path must refer to a file.");
  }

  if (path.extname(target).toLowerCase() !== ".md") {
    throw new Error("Manual path must refer to a Markdown file.");
  }

  return {
    absolutePath: target,
    path: normalizedPath,
  };
}

function getMarkdownTitle(text: string) {
  const heading = text.match(/^#\s+(.+)$/m);
  return heading?.[1]?.trim();
}

async function collectManualFiles(root: string, currentDirectory: string, results: ManualEntry[]) {
  const entries = await readdir(currentDirectory, { withFileTypes: true });

  for (const entry of entries) {
    const absolutePath = path.join(currentDirectory, entry.name);

    if (entry.isDirectory()) {
      await collectManualFiles(root, absolutePath, results);
      continue;
    }

    if (!entry.isFile() || path.extname(entry.name).toLowerCase() !== ".md") {
      continue;
    }

    const resolvedFilePath = await realpath(absolutePath);

    if (!isWithinRoot(root, resolvedFilePath)) {
      continue;
    }

    const relativePath = toManualPath(path.relative(root, resolvedFilePath));
    const text = await readFile(resolvedFilePath, "utf8");
    const title = getMarkdownTitle(text);

    results.push({
      absolutePath: resolvedFilePath,
      path: relativePath,
      readTool: {
        arguments: {
          path: relativePath,
        },
        name: "read_manual",
      },
      ...(title ? { title } : {}),
    });
  }
}

export async function listManuals(manualsDirectory: string): Promise<ManualListResult> {
  const root = await realpath(manualsDirectory);
  const manuals: ManualEntry[] = [];

  await collectManualFiles(root, root, manuals);
  manuals.sort((left, right) => left.path.localeCompare(right.path));

  return {
    manuals,
    manualsDirectory: root,
  };
}

export async function readManual(
  manualsDirectory: string,
  requestedPath: string,
  maxBytes = DEFAULT_MANUAL_READ_MAX_BYTES,
): Promise<ManualReadResult> {
  const { absolutePath, path: normalizedPath } = await resolveManualFilePath(
    manualsDirectory,
    requestedPath,
  );
  const text = await readFile(absolutePath, "utf8");
  const truncated = text.length > maxBytes;

  return {
    path: normalizedPath,
    text: truncated ? text.slice(0, maxBytes) : text,
    truncated,
  };
}
