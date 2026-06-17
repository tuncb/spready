import assert from "node:assert/strict";
import { mkdtemp, mkdir, realpath, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import {
  findManualsDirectory,
  listManuals,
  normalizeManualRequestPath,
  readManual,
  resolveManualFilePath,
} from "./mcp-manuals";

async function withTempDirectory(run: (directory: string) => Promise<void>) {
  const directory = await mkdtemp(path.join(os.tmpdir(), "spready-mcp-manuals-"));

  try {
    await run(directory);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
}

test("listManuals returns markdown manuals with filesystem and read_manual hints", async () => {
  await withTempDirectory(async (directory) => {
    await mkdir(path.join(directory, "nested"));
    await writeFile(path.join(directory, "formula.md"), "# Formula Manual\n\nDetails", "utf8");
    await writeFile(path.join(directory, "nested", "tables.md"), "# Tables\n\nDetails", "utf8");
    await writeFile(path.join(directory, "ignore.txt"), "# Ignore", "utf8");

    const result = await listManuals(directory);

    assert.equal(result.manualsDirectory, await realpath(directory));
    assert.deepEqual(
      result.manuals.map((manual) => ({
        path: manual.path,
        readTool: manual.readTool,
        title: manual.title,
      })),
      [
        {
          path: "formula.md",
          readTool: {
            arguments: {
              path: "formula.md",
            },
            name: "read_manual",
          },
          title: "Formula Manual",
        },
        {
          path: "nested/tables.md",
          readTool: {
            arguments: {
              path: "nested/tables.md",
            },
            name: "read_manual",
          },
          title: "Tables",
        },
      ],
    );
    assert.ok(result.manuals[0]?.absolutePath.endsWith(`${path.sep}formula.md`));
  });
});

test("readManual reads relative markdown paths and truncates large output", async () => {
  await withTempDirectory(async (directory) => {
    await mkdir(path.join(directory, "nested"));
    await writeFile(path.join(directory, "nested", "formula.md"), "abcdef", "utf8");

    assert.deepEqual(await readManual(directory, "nested\\formula.md", 3), {
      path: "nested/formula.md",
      text: "abc",
      truncated: true,
    });
    assert.deepEqual(await readManual(directory, "./nested/formula.md", 100), {
      path: "nested/formula.md",
      text: "abcdef",
      truncated: false,
    });
  });
});

test("manual path resolution rejects unsafe or unsupported paths", async () => {
  await withTempDirectory(async (directory) => {
    await writeFile(path.join(directory, "formula.md"), "# Formula", "utf8");
    await writeFile(path.join(directory, "notes.txt"), "notes", "utf8");

    assert.equal(normalizeManualRequestPath("docs\\formula.md"), "docs/formula.md");
    await assert.rejects(resolveManualFilePath(directory, "../outside.md"), /cannot escape/);
    await assert.rejects(
      resolveManualFilePath(directory, path.join(directory, "formula.md")),
      /relative/,
    );
    await assert.rejects(resolveManualFilePath(directory, "notes.txt"), /Markdown/);
  });
});

test("findManualsDirectory prefers explicit configuration and falls back to cwd manuals", async () => {
  await withTempDirectory(async (directory) => {
    const configured = path.join(directory, "configured");
    const cwd = path.join(directory, "project");

    await mkdir(configured);
    await mkdir(path.join(cwd, "manuals"), { recursive: true });

    assert.equal(await findManualsDirectory({ cwd, manualsDir: configured }), configured);
    assert.equal(
      await findManualsDirectory({
        cwd,
        executablePath: path.join(directory, "bin", "spready-mcp.exe"),
      }),
      path.join(cwd, "manuals"),
    );
  });
});
