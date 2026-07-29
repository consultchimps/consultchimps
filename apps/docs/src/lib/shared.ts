export const appName = "ConsultChimps";
export const docsRoute = "/docs";
export const docsImageRoute = "/og/docs";
export const docsContentRoute = "/llms.mdx/docs";

// Deployment path prefix for GitHub Pages project sites. Next.js prefixes
// <Link> and router navigation automatically; raw fetch targets and metadata
// URLs must prepend this manually.
export const basePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";

export const gitConfig = {
  user: "consultchimps",
  repo: "consultchimps",
  branch: "main",
};
