import { WorkspaceTool } from "@/components/workspace-tool";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Database in your browser",
  description:
    "Create or open a persistent local SQLite or DuckDB database, review workbook batches, record batch context, and export a portable copy.",
};

export default function Page() {
  return <WorkspaceTool />;
}
