import { resolve } from "node:path";
import { defineConfig } from "vite";

export default defineConfig({
  base: "/leanix-excel-addin/",
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html")
      }
    }
  }
});
