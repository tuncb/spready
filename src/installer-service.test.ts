import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import {
  extractSha256FromDigest,
  extractSha256FromText,
  getDefaultInstallDirectory,
  getInstalledExecutablePath,
  parseLatestReleaseResponse,
  parseVersionTag,
  selectSha256Asset,
  selectWindowsReleaseAsset,
  isVersionNewer,
} from "./installer-service";

test("parseVersionTag accepts semantic release tags", () => {
  assert.deepEqual(parseVersionTag("v1.2.3"), {
    major: 1,
    minor: 2,
    patch: 3,
  });
  assert.deepEqual(parseVersionTag("1.2.3-beta.1"), {
    major: 1,
    minor: 2,
    patch: 3,
  });
  assert.equal(parseVersionTag("release-1.2.3"), null);
});

test("isVersionNewer compares semantic versions", () => {
  assert.equal(
    isVersionNewer({ major: 1, minor: 3, patch: 0 }, { major: 1, minor: 2, patch: 9 }),
    true,
  );
  assert.equal(
    isVersionNewer({ major: 1, minor: 2, patch: 9 }, { major: 1, minor: 3, patch: 0 }),
    false,
  );
  assert.equal(
    isVersionNewer({ major: 1, minor: 2, patch: 3 }, { major: 1, minor: 2, patch: 3 }),
    false,
  );
});

test("extractSha256 helpers accept digest fields and checksum text", () => {
  const sha = "a".repeat(64);

  assert.equal(extractSha256FromDigest(`sha256:${sha}`), sha);
  assert.equal(extractSha256FromDigest(sha.toUpperCase()), sha);
  assert.equal(extractSha256FromText(`${sha}  spready.zip`), sha);
  assert.equal(extractSha256FromText("not a checksum"), null);
});

test("release parsing and asset selection target Windows bundles", () => {
  const release = parseLatestReleaseResponse(
    JSON.stringify({
      assets: [
        {
          browser_download_url: "https://example.test/spready-linux-x64-1.2.3.zip",
          name: "spready-linux-x64-1.2.3.zip",
        },
        {
          browser_download_url: "https://example.test/spready-windows-x64-1.2.3.zip",
          digest: `sha256:${"b".repeat(64)}`,
          name: "spready-windows-x64-1.2.3.zip",
        },
        {
          browser_download_url: "https://example.test/spready-windows-x64-1.2.3.zip.sha256",
          name: "spready-windows-x64-1.2.3.zip.sha256",
        },
      ],
      html_url: "https://example.test/releases/v1.2.3",
      tag_name: "v1.2.3",
    }),
  );

  const asset = selectWindowsReleaseAsset(release, "x64");

  assert.equal(asset?.name, "spready-windows-x64-1.2.3.zip");
  assert.equal(selectSha256Asset(release, asset?.name ?? "")?.name, `${asset?.name}.sha256`);
  assert.equal(release.html_url, "https://example.test/releases/v1.2.3");
});

test("install paths are derived from Windows local app data", () => {
  const installDirectory = getDefaultInstallDirectory(
    {
      LOCALAPPDATA: "C:\\Users\\person\\AppData\\Local",
    },
    "win32",
  );

  assert.equal(
    installDirectory,
    path.join("C:\\Users\\person\\AppData\\Local", "Programs", "Spready"),
  );
  assert.equal(
    getInstalledExecutablePath(installDirectory),
    path.join(installDirectory, "Spready.exe"),
  );
});

test("non-Windows install directory uses the user data area", () => {
  assert.equal(
    getDefaultInstallDirectory({}, "linux"),
    path.join(os.homedir(), ".local", "share", "Spready"),
  );
});
