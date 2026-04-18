import { mkdir, writeFile, chmod } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

import { build } from 'esbuild';

function parseArgs(argv) {
  const args = new Map();

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];

    if (!token.startsWith('--')) {
      continue;
    }

    const separatorIndex = token.indexOf('=');

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

    if (!value || value.startsWith('--')) {
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

function ensureSupportedNodeVersion() {
  const [majorText, minorText] = process.versions.node.split('.');
  const major = Number.parseInt(majorText, 10);
  const minor = Number.parseInt(minorText, 10);

  if (major > 25) {
    return;
  }

  if (major === 25 && minor >= 5) {
    return;
  }

  throw new Error(
    `Node ${process.versions.node} does not support --build-sea. Use Node 25.5 or newer.`,
  );
}

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');

async function main() {
  ensureSupportedNodeVersion();

  const args = parseArgs(process.argv.slice(2));
  const platform = requireArg(args, 'platform');
  const arch = requireArg(args, 'arch');
  const outputDir = path.join(repoRoot, 'out', 'mcp', `${platform}-${arch}`);
  const executableName = `spready-mcp${platform === 'win32' ? '.exe' : ''}`;
  const bundlePath = path.join(outputDir, 'spready-mcp.bundle.cjs');
  const seaConfigPath = path.join(outputDir, 'sea-config.json');
  const executablePath = path.join(outputDir, executableName);

  await mkdir(outputDir, { recursive: true });

  await build({
    bundle: true,
    entryPoints: [path.join(repoRoot, 'src', 'mcp-stdio.ts')],
    format: 'cjs',
    logLevel: 'info',
    outfile: bundlePath,
    platform: 'node',
    target: `node${process.versions.node}`,
  });

  const seaConfig = {
    disableExperimentalSEAWarning: true,
    main: bundlePath,
    output: executablePath,
    useCodeCache: false,
    useSnapshot: false,
  };

  await writeFile(seaConfigPath, `${JSON.stringify(seaConfig, null, 2)}\n`, 'utf8');

  const result = spawnSync(process.execPath, ['--build-sea', seaConfigPath], {
    cwd: repoRoot,
    stdio: 'inherit',
  });

  if (result.status !== 0) {
    throw new Error(`Node SEA build failed with exit code ${result.status ?? 'unknown'}.`);
  }

  if (platform !== 'win32') {
    await chmod(executablePath, 0o755);
  }

  console.log(`Built ${executablePath}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
