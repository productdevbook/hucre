import { resolve } from "node:path";
import { defineConfig } from "vite";
import { nitro } from "nitro/vite";

export default defineConfig({
  plugins: [
    nitro(),
  ],
  resolve: {
    alias: {
      hucre: resolve(__dirname, "../src/index.ts"),
      "hucre/xlsx": resolve(__dirname, "../src/xlsx.ts"),
      "hucre/csv": resolve(__dirname, "../src/csv.ts"),
      "hucre/ods": resolve(__dirname, "../src/ods.ts"),
    },
  },
});
