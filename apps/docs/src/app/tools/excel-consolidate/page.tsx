import { ExcelConsolidateTool } from "@/components/excel-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Consolidate Excel workbooks in your browser",
  description:
    "Stack the rows of every useful worksheet in several workbooks into one auditable table, even when columns arrive in different orders. It runs entirely in your browser tab.",
};

export default function Page() {
  return <ExcelConsolidateTool />;
}
