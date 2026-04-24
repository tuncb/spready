import { chmod, cp, mkdir, readdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

function parseArgs(argv) {
  const args = new Map();

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];

    if (!token.startsWith("--")) {
      continue;
    }

    const separatorIndex = token.indexOf("=");

    if (separatorIndex !== -1) {
      const key = token.slice(2, separatorIndex);
      const value = token.slice(separatorIndex + 1);

      if (!value) {
        throw new Error(`Missing value for --${key}.`);
      }

      args.set(key, value);
      continue;
    }

    const key = token.slice(2);
    const value = argv[index + 1];

    if (!value || value.startsWith("--")) {
      throw new Error(`Missing value for --${key}.`);
    }

    args.set(key, value);
    index += 1;
  }

  return args;
}

function requireArg(args, name) {
  const value = args.get(name);

  if (!value) {
    throw new Error(`Missing required argument --${name}.`);
  }

  return value;
}

function getPlatformSlug(platform) {
  switch (platform) {
    case "win32":
      return "windows";
    case "darwin":
      return "macos";
    case "linux":
      return "linux";
    default:
      throw new Error(`Unsupported platform "${platform}".`);
  }
}

function getExecutableName(productName, platform) {
  if (platform === "win32") {
    return `${productName}.exe`;
  }

  if (platform === "darwin") {
    return `${productName}.app`;
  }

  return productName;
}

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");

async function copyDirectoryContents(sourceDir, targetDir) {
  const entries = await readdir(sourceDir);

  for (const entry of entries) {
    await cp(path.join(sourceDir, entry), path.join(targetDir, entry), {
      recursive: true,
    });
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const platform = requireArg(args, "platform");
  const arch = requireArg(args, "arch");
  const version = requireArg(args, "version");
  const packageJson = JSON.parse(await readFile(path.join(repoRoot, "package.json"), "utf8"));
  const productName = packageJson.productName;

  if (typeof productName !== "string" || productName.length === 0) {
    throw new Error("package.json must define a productName.");
  }

  const platformSlug = getPlatformSlug(platform);
  const bundleSlug = `spready-${platformSlug}-${arch}-${version}`;
  const packagedAppDir = path.join(repoRoot, "out", `${productName}-${platform}-${arch}`);
  const mcpExecutableName = `spready-mcp${platform === "win32" ? ".exe" : ""}`;
  const mcpExecutablePath = path.join(
    repoRoot,
    "out",
    "mcp",
    `${platform}-${arch}`,
    mcpExecutableName,
  );
  const stagingDir = path.join(repoRoot, "out", "release", bundleSlug);
  const appTargetName = getExecutableName(productName, platform);
  const mcpConfigPath = path.join(stagingDir, "spready.mcp.json");
  const notesPath = path.join(stagingDir, "MCP-README.txt");
  const mcpCommand = `./${mcpExecutableName}`;
  const mcpConfig = {
    mcpServers: {
      spready: {
        command: mcpCommand,
      },
    },
  };
  const notes = [
    "1. Start Spready before connecting your harness.",
    `2. Import or copy the server entry from spready.mcp.json.`,
    `3. If your harness requires absolute command paths, replace ${mcpCommand} with the extracted path to ${mcpExecutableName}.`,
  ].join("\n");

  await stat(packagedAppDir);
  await stat(mcpExecutablePath);
  await rm(stagingDir, { force: true, recursive: true });
  await mkdir(stagingDir, { recursive: true });

  await copyDirectoryContents(packagedAppDir, stagingDir);
  await cp(mcpExecutablePath, path.join(stagingDir, mcpExecutableName), {
    recursive: false,
  });
  await writeFile(mcpConfigPath, `${JSON.stringify(mcpConfig, null, 2)}\n`, "utf8");
  await writeFile(notesPath, `${notes}\n`, "utf8");

  if (platform !== "win32") {
    await chmod(path.join(stagingDir, mcpExecutableName), 0o755);
  }

  console.log(`Prepared ${stagingDir}`);
  console.log(`Application bundle entry: ${appTargetName}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
