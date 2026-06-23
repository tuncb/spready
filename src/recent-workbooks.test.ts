import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import {
  parseRecentWorkbooksToml,
  RecentWorkbooksStore,
  serializeRecentWorkbooksToml,
} from "./recent-workbooks";

test("recent workbook TOML parser reads the generated files array", () => {
  const filePath = path.join("C:\\work", 'Budget "FY26".spready');
  const content = serializeRecentWorkbooksToml([
    {
      filePath,
      lastOpenedAt: "2026-06-23T10:00:00.000Z",
    },
  ]);

  assert.deepEqual(parseRecentWorkbooksToml(content), [
    {
      filePath,
      lastOpenedAt: "",
    },
  ]);
});

test("recent workbook TOML parser accepts structured workbook entries", () => {
  assert.deepEqual(
    parseRecentWorkbooksToml(`
[[workbooks]]
filePath = "C:\\\\work\\\\One.spready"
lastOpenedAt = "2026-06-23T10:00:00.000Z"

[[workbooks]]
filePath = "C:\\\\work\\\\Two.spready"
`),
    [
      {
        filePath: "C:\\work\\One.spready",
        lastOpenedAt: "2026-06-23T10:00:00.000Z",
      },
      {
        filePath: "C:\\work\\Two.spready",
        lastOpenedAt: "",
      },
    ],
  );
});

test("RecentWorkbooksStore persists, deduplicates, and limits recent workbooks", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spready-recents-"));
  const recentsPath = path.join(directory, "spready-recents.toml");
  let nowMs = Date.UTC(2026, 5, 23, 10, 0, 0);
  const store = new RecentWorkbooksStore({
    clock: () => {
      const date = new Date(nowMs);

      nowMs += 1000;
      return date;
    },
    filePath: recentsPath,
    maxEntries: 2,
  });

  await store.load();
  await store.add(path.join(directory, "One.spready"));
  await store.add(path.join(directory, "Two.spready"));
  const result = await store.add(path.join(directory, "One.spready"));

  assert.deepEqual(
    result.workbooks.map((entry) => path.basename(entry.filePath)),
    ["One.spready", "Two.spready"],
  );
  assert.deepEqual(
    result.workbooks.map((entry) => entry.lastOpenedAt),
    ["2026-06-23T10:00:02.000Z", "2026-06-23T10:00:01.000Z"],
  );

  const reloadedStore = new RecentWorkbooksStore({
    filePath: recentsPath,
    maxEntries: 2,
  });
  const reloadedResult = await reloadedStore.load();

  assert.deepEqual(
    reloadedResult.workbooks.map((entry) => path.basename(entry.filePath)),
    ["One.spready", "Two.spready"],
  );

  const removedResult = await reloadedStore.remove(path.join(directory, "Two.spready"));

  assert.deepEqual(
    removedResult.workbooks.map((entry) => path.basename(entry.filePath)),
    ["One.spready"],
  );

  const clearedResult = await reloadedStore.clear();

  assert.deepEqual(clearedResult.workbooks, []);
});
