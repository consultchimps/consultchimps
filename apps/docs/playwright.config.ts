/**
 * Playwright configuration for the in-browser tool pages.
 *
 * The suite is a smoke test of the statically exported site, not of the dev
 * server: it serves `out/` exactly as GitHub Pages would, so the lazily
 * imported PDF engine and the download plumbing are exercised through the
 * same bundles visitors receive.
 *
 * `pnpm --filter @consultchimps/docs... build` must run first. Build without
 * NEXT_PUBLIC_BASE_PATH so the routes are served from the site root.
 */
import { defineConfig, devices } from "@playwright/test";
import { existsSync } from "node:fs";
import path from "node:path";

const EXPORT_DIRECTORY = path.join(import.meta.dirname, "out");
const PORT = Number(process.env.PLAYWRIGHT_PORT ?? "4321");
const BASE_URL = `http://127.0.0.1:${PORT}`;

if (!existsSync(path.join(EXPORT_DIRECTORY, "index.html"))) {
  throw new Error(
    "apps/docs/out is missing or empty. Run `pnpm --filter @consultchimps/docs... build` (without NEXT_PUBLIC_BASE_PATH) before `pnpm --filter @consultchimps/docs e2e`.",
  );
}

export default defineConfig({
  testDir: "./e2e",
  fullyParallel: true,
  // A stray `test.only` should fail the pull request instead of quietly
  // shrinking the suite.
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 2 : 0,
  // One worker in CI keeps the report deterministic; locally the suite shares
  // the machine with whatever else is running.
  workers: process.env.CI ? 1 : "50%",
  reporter: process.env.CI
    ? [["list"], ["html", { open: "never" }]]
    : [["list"]],
  timeout: 60_000,
  expect: { timeout: 15_000 },
  use: {
    baseURL: BASE_URL,
    trace: "on-first-retry",
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
  ],
  webServer: {
    // `serve` maps the extensionless routes of the static export onto their
    // .html files; a plain file server would 404 on /tools/pdf-split. It is
    // started through its module path because Playwright does not add
    // node_modules/.bin to PATH for webServer commands.
    command: `node ./node_modules/serve/build/main.js out --listen ${PORT} --no-clipboard --no-port-switching`,
    url: `${BASE_URL}/tools`,
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
