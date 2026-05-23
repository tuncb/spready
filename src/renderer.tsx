let didLogReactMounted = false;

function getErrorDetail(error: unknown) {
  return error instanceof Error ? error.message : "unknown error";
}

async function bootstrapRenderer() {
  window.appShell.logStartupTiming("renderer-entry");
  window.appShell.logStartupTiming("renderer-bootstrap-entry");

  window.appShell.logStartupTiming("glide-css-import-start");
  await import("@glideapps/glide-data-grid/dist/index.css");
  window.appShell.logStartupTiming("glide-css-import-done");

  window.appShell.logStartupTiming("app-css-import-start");
  await import("./index.css");
  window.appShell.logStartupTiming("app-css-import-done");

  window.appShell.logStartupTiming("react-import-start");
  const { StrictMode, createElement, useEffect } = await import("react");
  window.appShell.logStartupTiming("react-import-done");

  window.appShell.logStartupTiming("react-dom-client-import-start");
  const { createRoot } = await import("react-dom/client");
  window.appShell.logStartupTiming("react-dom-client-import-done");

  window.appShell.logStartupTiming("app-module-import-start");
  const { default: App } = await import("./App");
  window.appShell.logStartupTiming("app-module-import-done");

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

  const app = document.querySelector<HTMLDivElement>("#app");

  if (!app) {
    throw new Error("App root was not found");
  }

  window.appShell.logStartupTiming("react-render-start");

  createRoot(app).render(
    createElement(StrictMode, null, createElement(StartupTimingProbe), createElement(App)),
  );

  window.appShell.logStartupTiming("react-render-dispatched");
}

void bootstrapRenderer().catch((error: unknown) => {
  window.appShell.logStartupTiming("renderer-bootstrap-failed", getErrorDetail(error));
});
