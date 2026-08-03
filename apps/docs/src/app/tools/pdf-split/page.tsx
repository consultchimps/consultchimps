import { PdfSplitTool } from "@/components/pdf-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Split a PDF in your browser",
  description:
    "Write every page of a PDF to its own file without uploading anything. The split runs entirely in your browser tab.",
};

export default function Page() {
  return <PdfSplitTool />;
}
