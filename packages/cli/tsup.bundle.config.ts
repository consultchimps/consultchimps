import { readFileSync } from "node:fs";
import { defineConfig } from "tsup";

const { version } = JSON.parse(
  readFileSync(new URL("./package.json", import.meta.url), "utf8"),
) as { version: string };

// Builds the standalone single-file CLI attached to GitHub releases. Every
// runtime dependency is compiled in so `node consultchimps.mjs` performs no
// package resolution at all — only `node:` builtins remain external — for
// environments that have Node.js (22+) but no npm access.
export default defineConfig({
  entry: { consultchimps: "src/index.ts" },
  format: "esm",
  outDir: "dist-bundle",
  noExternal: [/.*/],
  splitting: false,
  clean: true,
  outExtension: () => ({ js: ".mjs" }),
  define: {
    CONSULTCHIMPS_BUNDLED_VERSION: JSON.stringify(version),
  },
  banner: {
    // Bundled CommonJS dependencies require() node builtins at runtime;
    // ESM output has no require in scope unless we create one.
    js: 'import { createRequire as __createRequire } from "node:module"; const require = __createRequire(import.meta.url);',
  },
});
