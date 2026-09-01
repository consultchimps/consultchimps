import { ExcelInspectTool } from "@/components/excel-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Inspect an Excel workbook in your browser",
  description:
    "Report a workbook's worksheets, hidden tabs, header rows, Excel Tables, named ranges, and sample column values before you operate on it. The workbook is read in your browser tab and never uploaded.",
};

export default function Page() {
  return <ExcelInspectTool />;
}
