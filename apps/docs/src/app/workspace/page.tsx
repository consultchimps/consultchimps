import { WorkspaceTool } from "@/components/workspace-tool";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Data workspace in your browser",
  description:
    "Start a local data workspace, import an Excel worksheet or a .csv file into it, and save it back to one file. The workspace is an in-memory database that stays in your browser tab.",
};

export default function WorkspacePage() {
  return <WorkspaceTool />;
}
