import { PptxInspectTool } from "@/components/pptx-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Inspect a PowerPoint template in your browser",
  description:
    "Report every placeholder a template slide expects, with occurrence counts, before you populate it. The template is read in your browser tab and never uploaded.",
};

export default function Page() {
  return <PptxInspectTool />;
}
