import {
  FileSearch,
  FileStack,
  GitMerge,
  Presentation,
  ScanLine,
  SplitSquareVertical,
  TableProperties,
  type LucideIcon,
} from "lucide-react";

/**
 * Single source of truth for every operation the toolkit ships, per ADR 0001
 * (docs/adr/0001-feature-registry-and-drift-checks.md): one entry per
 * operation, each declaring the status of its CLI, library, and browser
 * surfaces. Landing cards, online-tool tabs, and guide-page "Try … online"
 * buttons all derive from these statuses; `planned` and `none` render
 * nothing user-visible, so a page can never offer a capability that does
 * not exist.
 */

export type SurfaceStatus = "works" | "planned" | "none";

/**
 * The browser surface carries its route only when it works, so a card, tab,
 * or button can never link to a browser page that does not exist.
 */
export type BrowserSurface =
  | { readonly status: "works"; readonly href: string }
  | { readonly status: "planned" | "none" };

export interface ToolSurfaces {
  readonly cli: SurfaceStatus;
  readonly library: SurfaceStatus;
  readonly browser: BrowserSurface;
}

export interface ConsultTool {
  readonly slug: string;
  readonly title: string;
  /** Short name shown in the online-tools sub-bar and "Try … online" buttons. */
  readonly tabLabel: string;
  readonly description: string;
  readonly docHref: string;
  readonly surfaces: ToolSurfaces;
  readonly icon: LucideIcon;
}

/** A tool whose browser surface works — the only kind that renders browser UI. */
export interface BrowserTool extends ConsultTool {
  readonly surfaces: ToolSurfaces & {
    readonly browser: Extract<BrowserSurface, { status: "works" }>;
  };
}

export function isBrowserTool(tool: ConsultTool): tool is BrowserTool {
  return tool.surfaces.browser.status === "works";
}

export const TOOLS: readonly ConsultTool[] = [
  {
    slug: "spreadsheet-consolidate",
    title: "Consolidate spreadsheets",
    tabLabel: "Consolidate",
    description:
      "Stack rows from every useful worksheet into one auditable table, even when columns arrive in different orders.",
    docHref: "/docs/tools/spreadsheet-consolidate",
    surfaces: {
      cli: "works",
      library: "works",
      browser: { status: "none" },
    },
    icon: TableProperties,
  },
  {
    slug: "workbook-merge",
    title: "Merge workbook tabs",
    tabLabel: "Merge tabs",
    description:
      "Copy every source worksheet into one workbook as its own separate tab, never stacking rows into one sheet.",
    docHref: "/docs/tools/workbook-merge",
    surfaces: {
      cli: "works",
      library: "works",
      browser: { status: "works", href: "/tools/excel-merge" },
    },
    icon: FileStack,
  },
  {
    slug: "spreadsheet-split",
    title: "Split spreadsheets",
    tabLabel: "Split Excel",
    description:
      "Create one Excel workbook per distinct value, each a complete copy of the original, while keeping source workbooks unchanged.",
    docHref: "/docs/tools/spreadsheet-split",
    surfaces: {
      cli: "works",
      library: "works",
      browser: { status: "works", href: "/tools/excel-split" },
    },
    icon: SplitSquareVertical,
  },
  {
    slug: "powerpoint-populate",
    title: "Populate PowerPoint templates",
    tabLabel: "PowerPoint",
    description:
      "Turn a designed template slide and Excel records into a review-ready presentation, entirely locally.",
    docHref: "/docs/tools/powerpoint-populate",
    surfaces: {
      cli: "works",
      library: "works",
      browser: { status: "none" },
    },
    icon: Presentation,
  },
  {
    slug: "pdf-split",
    title: "Split PDF pages",
    tabLabel: "Split PDF",
    description:
      "Turn a long PDF into predictable, zero-padded page files without sending the document anywhere.",
    docHref: "/docs/tools/pdf-split",
    surfaces: {
      cli: "works",
      library: "works",
      browser: { status: "works", href: "/tools/pdf-split" },
    },
    icon: ScanLine,
  },
  {
    slug: "pdf-merge",
    title: "Merge PDF packs",
    tabLabel: "Merge PDFs",
    description:
      "Assemble source PDFs in resolved order and preserve every page in one clean deliverable.",
    docHref: "/docs/tools/pdf-merge",
    surfaces: {
      cli: "works",
      library: "works",
      browser: { status: "works", href: "/tools/pdf-merge" },
    },
    icon: GitMerge,
  },
  {
    slug: "powerpoint-inspect-template",
    title: "Inspect PowerPoint templates",
    tabLabel: "Inspect template",
    description:
      "Report every placeholder a template slide expects, with occurrence counts and malformed braces, before you populate it.",
    docHref: "/docs/tools/powerpoint-populate#inspect-the-template",
    surfaces: {
      cli: "works",
      library: "works",
      browser: { status: "none" },
    },
    icon: FileSearch,
  },
] as const;

export const BROWSER_TOOLS: readonly BrowserTool[] =
  TOOLS.filter(isBrowserTool);

/**
 * Find the tool whose working browser surface a docs page should offer.
 * Keyed off browser status, not entry order: when several operations share
 * one guide page (populate and template inspection both live on the
 * PowerPoint guide), the button belongs to the operation that actually runs
 * in the browser — or to none at all. Fragments are stripped so an anchored
 * docHref still resolves to its page.
 */
export function findToolByDocUrl(url: string): BrowserTool | undefined {
  return BROWSER_TOOLS.find((tool) => tool.docHref.split("#")[0] === url);
}
