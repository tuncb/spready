import "@glideapps/glide-data-grid/dist/index.css";
import "./index.css";

import { StrictMode, useEffect } from "react";
import { createRoot } from "react-dom/client";

import App from "./App";

let didLogReactMounted = false;

function StartupTimingProbe() {
  useEffect(() => {
    if (didLogReactMounted) {
      return;
    }

    didLogReactMounted = true;
    window.appShell.logStartupTiming("react-mounted");
    requestAnimationFrame(() => {
      window.appShell.logStartupTiming("first-animation-frame");
      requestAnimationFrame(() => {
        window.appShell.logStartupTiming("second-animation-frame");
      });
    });
  }, []);

  return null;
}

window.appShell.logStartupTiming("renderer-entry");

const app = document.querySelector<HTMLDivElement>("#app");

if (!app) {
  throw new Error("App root was not found");
}

window.appShell.logStartupTiming("react-render-start");

createRoot(app).render(
  <StrictMode>
    <StartupTimingProbe />
    <App />
  </StrictMode>,
);

window.appShell.logStartupTiming("react-render-dispatched");
