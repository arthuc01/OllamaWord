import { defineConfig } from "vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";

// Office desktop add-ins should be served over HTTPS during development.
// office-addin-dev-certs creates/trusts a local development certificate.
async function getHttpsOptions() {
  return await getHttpsServerOptions();
}

export default defineConfig(async () => ({
  root: ".",
  server: {
    host: "localhost",
    port: 3000,
    https: await getHttpsOptions()
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: "src/taskpane/taskpane.html",
        commands: "src/commands/commands.html"
      }
    }
  }
}));
