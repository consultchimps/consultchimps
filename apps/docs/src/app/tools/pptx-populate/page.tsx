import { PptxPopulateTool } from "@/components/pptx-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Populate a PowerPoint template in your browser",
  description:
    "Turn a designed template slide and a workbook of records into a presentation without uploading anything. The population runs entirely in your browser tab.",
};

export default function Page() {
  return <PptxPopulateTool />;
}
