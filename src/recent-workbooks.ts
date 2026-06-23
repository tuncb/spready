import { promises as fs } from "node:fs";
import path from "node:path";

export const DEFAULT_RECENT_WORKBOOKS_FILE_NAME = "spready-recents.toml";
export const DEFAULT_MAX_RECENT_WORKBOOKS = 10;

export interface RecentWorkbookEntry {
  filePath: string;
  lastOpenedAt: string;
}

export interface RecentWorkbooksResult {
  filePath: string;
  workbooks: RecentWorkbookEntry[];
}

type RecentWorkbooksStoreOptions = {
  clock?: () => Date;
  filePath: string;
  maxEntries?: number;
};

export class RecentWorkbooksStore {
  #clock: () => Date;
  #entries: RecentWorkbookEntry[] = [];
  #filePath: string;
  #maxEntries: number;

  constructor(options: RecentWorkbooksStoreOptions) {
    this.#clock = options.clock ?? (() => new Date());
    this.#filePath = path.resolve(options.filePath);
    this.#maxEntries = options.maxEntries ?? DEFAULT_MAX_RECENT_WORKBOOKS;
  }

  get filePath() {
    return this.#filePath;
  }

  async load(): Promise<RecentWorkbooksResult> {
    try {
      const content = await fs.readFile(this.#filePath, "utf8");

      this.#entries = normalizeRecentWorkbookEntries(
        parseRecentWorkbooksToml(content),
        this.#maxEntries,
      );
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
        throw error;
      }

      this.#entries = [];
    }

    return this.getRecentWorkbooks();
  }

  getRecentWorkbooks(): RecentWorkbooksResult {
    return {
      filePath: this.#filePath,
      workbooks: this.#entries.map((entry) => ({ ...entry })),
    };
  }

  async add(filePath: string): Promise<RecentWorkbooksResult> {
    const resolvedFilePath = path.resolve(filePath);
    const key = getRecentWorkbookKey(resolvedFilePath);
    const nextEntries = this.#entries.filter(
      (entry) => getRecentWorkbookKey(entry.filePath) !== key,
    );

    nextEntries.unshift({
      filePath: resolvedFilePath,
      lastOpenedAt: this.#clock().toISOString(),
    });

    this.#entries = normalizeRecentWorkbookEntries(nextEntries, this.#maxEntries);
    await this.#save();

    return this.getRecentWorkbooks();
  }

  async remove(filePath: string): Promise<RecentWorkbooksResult> {
    const resolvedFilePath = path.resolve(filePath);
    const key = getRecentWorkbookKey(resolvedFilePath);

    this.#entries = this.#entries.filter((entry) => getRecentWorkbookKey(entry.filePath) !== key);
    await this.#save();

    return this.getRecentWorkbooks();
  }

  async clear(): Promise<RecentWorkbooksResult> {
    this.#entries = [];
    await this.#save();

    return this.getRecentWorkbooks();
  }

  async #save() {
    await fs.mkdir(path.dirname(this.#filePath), { recursive: true });
    await fs.writeFile(this.#filePath, serializeRecentWorkbooksToml(this.#entries), "utf8");
  }
}

export function getDefaultRecentWorkbooksFilePath(executablePath: string) {
  return path.join(path.dirname(executablePath), DEFAULT_RECENT_WORKBOOKS_FILE_NAME);
}

export function parseRecentWorkbooksToml(content: string): RecentWorkbookEntry[] {
  const entries: RecentWorkbookEntry[] = [];
  const filesArrayMatch = /^\s*files\s*=\s*\[([\s\S]*?)^\s*\]/mu.exec(content);

  if (filesArrayMatch) {
    for (const value of parseTomlStringValues(filesArrayMatch[1])) {
      entries.push({
        filePath: value,
        lastOpenedAt: "",
      });
    }
  }

  const tableContents = content.split(/^\s*\[\[workbooks\]\]\s*$/gmu).slice(1);

  for (const tableContent of tableContents) {
    const filePath = parseTomlStringProperty(tableContent, "filePath");

    if (!filePath) {
      continue;
    }

    entries.push({
      filePath,
      lastOpenedAt: parseTomlStringProperty(tableContent, "lastOpenedAt") ?? "",
    });
  }

  return entries;
}

export function serializeRecentWorkbooksToml(entries: readonly RecentWorkbookEntry[]) {
  const lines = [
    "# Spready recent workbooks",
    "",
    "files = [",
    ...entries.map((entry) => `  ${formatTomlBasicString(entry.filePath)},`),
    "]",
    "",
  ];

  return `${lines.join("\n")}`;
}

function normalizeRecentWorkbookEntries(
  entries: readonly RecentWorkbookEntry[],
  maxEntries: number,
): RecentWorkbookEntry[] {
  const normalizedEntries: RecentWorkbookEntry[] = [];
  const seenKeys = new Set<string>();

  for (const entry of entries) {
    if (typeof entry.filePath !== "string" || entry.filePath.trim().length === 0) {
      continue;
    }

    const resolvedFilePath = path.resolve(entry.filePath);
    const key = getRecentWorkbookKey(resolvedFilePath);

    if (seenKeys.has(key)) {
      continue;
    }

    seenKeys.add(key);
    normalizedEntries.push({
      filePath: resolvedFilePath,
      lastOpenedAt: isValidIsoDate(entry.lastOpenedAt) ? entry.lastOpenedAt : "",
    });

    if (normalizedEntries.length >= maxEntries) {
      break;
    }
  }

  return normalizedEntries;
}

function getRecentWorkbookKey(filePath: string) {
  return process.platform === "win32" ? filePath.toLowerCase() : filePath;
}

function parseTomlStringProperty(content: string, propertyName: string): string | undefined {
  const match = new RegExp(`^\\s*${propertyName}\\s*=\\s*("(?:\\\\.|[^"\\\\])*")`, "mu").exec(
    content,
  );

  if (!match) {
    return undefined;
  }

  return parseTomlBasicString(match[1]);
}

function parseTomlStringValues(content: string): string[] {
  return [...content.matchAll(/"(?:\\.|[^"\\])*"/gu)].map((match) =>
    parseTomlBasicString(match[0]),
  );
}

function parseTomlBasicString(value: string): string {
  return JSON.parse(value) as string;
}

function formatTomlBasicString(value: string): string {
  return JSON.stringify(value);
}

function isValidIsoDate(value: string) {
  return value.length > 0 && !Number.isNaN(Date.parse(value));
}
