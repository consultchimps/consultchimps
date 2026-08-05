import {
  FileStack,
  GitMerge,
  Presentation,
  ScanLine,
  SplitSquareVertical,
  TableProperties,
  type LucideIcon,
} from "lucide-react";

/**
 * Single source of truth for every ConsultChimps tool surfaced on the site.
 *
 * `browserHref` doubles as a feature flag: when a tool has an in-browser
 * version, setting it lights up the "Try it online" button on the tool's
 * guide page, adds a tab to the online-tools sub-bar, and links the landing
 * card — all without touching those surfaces.
 */
export interface ConsultTool {
  readonly slug: string;
  readonly title: string;
  readonly tabLabel: string;
  readonly description: string;
  readonly docHref: string;
  readonly browserHref?: string;
  readonly icon: LucideIcon;
}

export const TOOLS: readonly ConsultTool[] = [
  {
    slug: "spreadsheet-consolidate",
    title: "Consolidate spreadsheets",
    tabLabel: "Consolidate",
    description:
      "Union every useful worksheet into one auditable table, even when columns arrive in different orders.",
    docHref: "/docs/tools/spreadsheets",
    icon: TableProperties,
  },
  {
    slug: "workbook-merge",
    title: "Merge workbook tabs",
    tabLabel: "Merge tabs",
    description:
      "Copy every source worksheet into one workbook while retaining separate tabs and source visibility.",
    docHref: "/docs/tools/spreadsheets#merge-complete-workbooks",
    browserHref: "/tools/excel-merge",
    icon: FileStack,
  },
  {
    slug: "spreadsheet-split",
    title: "Split spreadsheets",
    tabLabel: "Split Excel",
    description:
      "Create one focused Excel workbook per distinct value while keeping source workbooks unchanged.",
    docHref: "/docs/tools/spreadsheet-split",
    browserHref: "/tools/excel-split",
    icon: SplitSquareVertical,
  },
  {
    slug: "powerpoint-populate",
    title: "Populate PowerPoint templates",
    tabLabel: "PowerPoint",
    description:
      "Turn a designed template slide and Excel records into a review-ready presentation, entirely locally.",
    docHref: "/docs/tools/powerpoint-populate",
    icon: Presentation,
  },
  {
    slug: "pdf-split",
    title: "Split PDF pages",
    tabLabel: "Split PDF",
    description:
      "Turn a long PDF into predictable, zero-padded page files without sending the document anywhere.",
    docHref: "/docs/tools/pdf-split",
    browserHref: "/tools/pdf-split",
    icon: ScanLine,
  },
  {
    slug: "pdf-merge",
    title: "Merge PDF packs",
    tabLabel: "Merge PDFs",
    description:
      "Assemble source PDFs in resolved order and preserve every page in one clean deliverable.",
    docHref: "/docs/tools/pdf-merge",
    browserHref: "/tools/pdf-merge",
    icon: GitMerge,
  },
] as const;

export const BROWSER_TOOLS: readonly ConsultTool[] = TOOLS.filter(
  (tool) => tool.browserHref !== undefined,
);

export function findToolByDocUrl(url: string): ConsultTool | undefined {
  return TOOLS.find(
    (tool) =>
      tool.docHref.split("#")[0] === url && tool.browserHref !== undefined,
  );
}
