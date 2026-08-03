import { PdfMergeTool } from "@/components/pdf-tools";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Merge PDFs in your browser",
  description:
    "Combine several PDFs into one document in the order you choose. The merge runs entirely in your browser tab.",
};

export default function Page() {
  return <PdfMergeTool />;
}
