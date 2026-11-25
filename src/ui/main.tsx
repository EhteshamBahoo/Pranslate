import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./styles.css";

function diagnostics(info: { host: unknown; platform: unknown }): string {
  const ok = (v: number) => {
    try { return Office.context.requirements.isSetSupported("PowerPointApi", String(v)); } catch { return "n/a"; }
  };
  return [
    `host: ${String(info.host)}  platform: ${String(info.platform)}`,
    `typeof PowerPoint: ${typeof PowerPoint}`,
    `PowerPointApi 1.4: ${ok(1.4)}  1.5: ${ok(1.5)}  1.8: ${ok(1.8)}`,
  ].join("\n");
}

if (typeof Office === "undefined") {
  document.getElementById("root")!.textContent =
    "Office.js did not load (blocked or offline). Check access to appsforoffice.microsoft.com.";
} else {
  Office.onReady((info) => {
    createRoot(document.getElementById("root")!).render(<App diag={diagnostics(info)} />);
  });
}
