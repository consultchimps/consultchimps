import { ExcelMergeTool } from "@/components/excel-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Merge Excel workbooks in your browser",
  description:
    "Combine every worksheet of several workbooks into one, keeping separate tabs and formatting. The merge runs entirely in your browser tab.",
};

export default function Page() {
  return <ExcelMergeTool />;
}
