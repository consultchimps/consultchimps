import { createMDX } from "fumadocs-mdx/next";
import type { NextConfig } from "next";

const config: NextConfig = {
  allowedDevOrigins: ["127.0.0.1"],
  reactStrictMode: true,
  // Static export for GitHub Pages. NEXT_PUBLIC_BASE_PATH is set by the
  // deploy workflow ("/consultchimps") and empty in local development.
  output: "export",
  basePath: process.env.NEXT_PUBLIC_BASE_PATH ?? "",
  images: {
    unoptimized: true,
  },
};

export default createMDX()(config);
