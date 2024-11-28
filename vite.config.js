import { defineConfig } from "vite";
import topLevelAwait from "vite-plugin-top-level-await";

export default defineConfig({
  root: "./examples",
  plugins: [topLevelAwait()],
  // test: {
  //   include: [],
  // },
});
