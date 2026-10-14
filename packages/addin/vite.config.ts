import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

/**
 * Dev server runs on https://localhost:3000 using the Office add-in dev
 * certificate (`npm run certs -w @loai/addin` once). Production builds are
 * served by the companion at https://localhost:47615.
 */
async function httpsOptions() {
  try {
    const certs = await import("office-addin-dev-certs");
    return await certs.getHttpsServerOptions();
  } catch {
    console.warn("[loai] Dev certificate not found; serving over HTTP. Run `npm run certs -w @loai/addin` for Office.");
    return undefined;
  }
}

export default defineConfig(async ({ command }) => ({
  root: ".",
  resolve: {
    alias: { "@loai/core": fileURLToPath(new URL("../core/src/index.ts", import.meta.url)) },
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: true,
    rollupOptions: { input: { taskpane: fileURLToPath(new URL("./taskpane.html", import.meta.url)) } },
  },
  server: {
    port: 3000,
    strictPort: true,
    https: command === "serve" ? await httpsOptions() : undefined,
  },
}));
