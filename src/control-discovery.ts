import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import type { ControlServerInfo } from './workbook-core';

export interface DiscoveredControlInfo extends ControlServerInfo {
  appName: string;
  pid: number;
  updatedAt: string;
}

export const CONTROL_DISCOVERY_FILE_PATH = path.join(os.tmpdir(), 'spready-control.json');

export async function clearDiscoveredControlInfo(expectedPid = process.pid) {
  try {
    const current = await readDiscoveredControlInfo();

    if (!current || current.pid !== expectedPid) {
      return;
    }

    await fs.unlink(CONTROL_DISCOVERY_FILE_PATH);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') {
      throw error;
    }
  }
}

export async function readDiscoveredControlInfo(): Promise<DiscoveredControlInfo | null> {
  try {
    const content = await fs.readFile(CONTROL_DISCOVERY_FILE_PATH, 'utf8');
    const parsed = JSON.parse(content) as Partial<DiscoveredControlInfo>;

    if (
      typeof parsed.appName !== 'string' ||
      typeof parsed.host !== 'string' ||
      typeof parsed.pid !== 'number' ||
      typeof parsed.port !== 'number' ||
      parsed.protocol !== 'jsonl' ||
      typeof parsed.updatedAt !== 'string'
    ) {
      return null;
    }

    return parsed as DiscoveredControlInfo;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
      return null;
    }

    throw error;
  }
}

export async function writeDiscoveredControlInfo(
  appName: string,
  controlInfo: ControlServerInfo,
  pid = process.pid,
) {
  const discoveredControlInfo: DiscoveredControlInfo = {
    appName,
    host: controlInfo.host,
    pid,
    port: controlInfo.port,
    protocol: controlInfo.protocol,
    updatedAt: new Date().toISOString(),
  };

  await fs.writeFile(
    CONTROL_DISCOVERY_FILE_PATH,
    JSON.stringify(discoveredControlInfo, null, 2),
    'utf8',
  );
}
