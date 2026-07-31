import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["packages/**/*.test.ts"],
    testTimeout: 15_000,
    coverage: {
      provider: "v8",
      include: ["packages/*/src/**"],
      // The CLI is tested by executing its built binary in a subprocess, so
      // source coverage cannot be attributed to packages/cli/src.
      exclude: ["packages/cli/src/**"],
      reporter: ["text-summary", "html"],
      // Regression floors slightly below current coverage, not aspirations.
      thresholds: {
        branches: 70,
        functions: 90,
        lines: 85,
        statements: 85,
      },
    },
  },
});
