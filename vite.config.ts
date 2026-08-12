import { resolve } from "node:path";
import { defineConfig } from "vite";

export default defineConfig({
  base: "/leanix-excel-addin/",
  build: {
    rollupOptions: {
      input: {
        bridgeRedirect: resolve(__dirname, "bridge-redirect.html"),
        bridgeUrlPrompt: resolve(__dirname, "bridge-url-prompt.html"),
        taskpane: resolve(__dirname, "taskpane.html")
      }
    }
  }
});
