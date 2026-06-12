import type { ForgeConfig } from "@electron-forge/shared-types";
import { MakerSquirrel } from "@electron-forge/maker-squirrel";
import { VitePlugin } from "@electron-forge/plugin-vite";
import path from "node:path";

const iconBasePath = path.resolve(__dirname, "assets", "spready");
const windowsIconPath = `${iconBasePath}.ico`;

const config: ForgeConfig = {
  packagerConfig: {
    asar: true,
    icon: iconBasePath,
  },
  rebuildConfig: {},
  makers: [
    new MakerSquirrel({
      authors: "Spready",
      setupIcon: windowsIconPath,
    }),
  ],
  plugins: [
    new VitePlugin({
      build: [
        {
          entry: "src/main.ts",
          config: "vite.main.config.ts",
          target: "main",
        },
        {
          entry: "src/preload.ts",
          config: "vite.preload.config.ts",
          target: "preload",
        },
      ],
      renderer: [
        {
          name: "main_window",
          config: "vite.renderer.config.ts",
        },
      ],
    }),
  ],
};

export default config;
