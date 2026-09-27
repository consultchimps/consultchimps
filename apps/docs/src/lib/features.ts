/**
 * Site feature flags. A disabled feature leaves the built site entirely: its
 * guide pages (dropped by the loader in source.ts), its `<Feature name="…">`
 * sections in MDX (removed at compile time by remarkFeatureFlags), and its
 * registry entries (filtered out of SITE_TOOLS). The drift checks read the MDX
 * source and the full registry, so hidden content stays checked.
 */
export const FEATURES = {
  // Paused while the database tool and the `consultchimps db` commands are
  // reworked; the commands are not in a published CLI release yet.
  database: false,
} as const satisfies Record<string, boolean>;

export type Feature = keyof typeof FEATURES;

/** Guide pages that belong to a feature, as paths under content/docs. */
export const FEATURE_PAGES: Readonly<Record<Feature, readonly string[]>> = {
  database: ["tools/data-workspace.mdx"],
};

/** The kinds of tool the site describes itself as offering. */
export const TOOL_KINDS = FEATURES.database
  ? "spreadsheet, database, PowerPoint, and PDF"
  : "spreadsheet, PowerPoint, and PDF";

export function isFeatureEnabled(feature: Feature | undefined): boolean {
  return feature === undefined || FEATURES[feature];
}

export function isFeature(name: string): name is Feature {
  return Object.hasOwn(FEATURES, name);
}

/** Whether a content/docs page belongs to a disabled feature. */
export function isHiddenPage(contentPath: string): boolean {
  const normalized = contentPath.replaceAll("\\", "/");
  return Object.entries(FEATURE_PAGES).some(
    ([feature, pages]) =>
      !FEATURES[feature as Feature] && pages.includes(normalized),
  );
}
