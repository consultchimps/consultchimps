import { ExcelSplitTool } from "@/components/excel-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Split an Excel workbook in your browser",
  description:
    "Create one workbook per distinct value in a column without uploading anything. The split runs entirely in your browser tab.",
};

export default function Page() {
  return <ExcelSplitTool />;
}
