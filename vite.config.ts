import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";

// Office add-ins must be served over HTTPS. Dev certs are created with:
//   npx office-addin-dev-certs install
// They land in ~/.office-addin-dev-certs/. If absent, Vite falls back to HTTP
// (the add-in will NOT load in PowerPoint, but unit tests and typecheck still work).
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const keyPath = path.join(certDir, "localhost.key");
const certPath = path.join(certDir, "localhost.crt");
const https =
  fs.existsSync(keyPath) && fs.existsSync(certPath)
    ? { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) }
    : undefined;

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    https,
    // Dev-only: same-origin proxy to a local Ollama so the HTTPS task pane avoids CORS and mixed-content blocks.
    proxy: {
      "/ollama": {
        target: "http://localhost:11434",
        changeOrigin: true,
        rewrite: (p) => p.replace(/^\/ollama/, ""),
        configure: (proxy) => {
          // Ollama rejects unknown Origins; drop the header on the proxied request.
          proxy.on("proxyReq", (req) => req.removeHeader("origin"));
        },
      },
    },
  },
  test: { include: ["src/**/*.test.ts"] },
});
