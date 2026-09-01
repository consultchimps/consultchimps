"use client";

import {
  ChosenFile,
  FilePicker,
  ReadingFile,
  ResultsPanel,
  RunControls,
  sectionClass,
  ToolShell,
  useFileSelection,
  useOperationRun,
} from "@/components/tool-kit";
import { WORKBOOK_FILES } from "@/lib/accepted-files";
import type { ByteOperationTask } from "@/lib/operation-tasks";
import { useCallback } from "react";

export function XlsxUnprotectTool() {
  const selection = useFileSelection(
    WORKBOOK_FILES.accepts,
    WORKBOOK_FILES.description,
  );
  const runState = useOperationRun();
  const start = useCallback(() => {
    if (!selection.file) return;
    const outputName = selection.file.name.replace(
      /\.(xlsx|xlsm)$/iu,
      "-unprotected.$1",
    );
    void runState.run({
      kind: "xlsx.unprotect",
      input: selection.file,
      outputName,
    } satisfies ByteOperationTask);
  }, [runState, selection.file]);
  return (
    <ToolShell
      description="Remove ordinary worksheet and workbook-structure protection locally. Nothing is uploaded, and your original workbook is untouched."
      guideHref="/docs/tools/excel-unprotect"
      guideLabel="Read the unprotect guide"
      kicker="Online tool · Excel unprotect"
      title="Unprotect an Excel workbook"
    >
      <section className={sectionClass}>
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Choose a workbook
        </h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          Supported: .xlsx and .xlsm. Encrypted files that require a password to
          open are not supported.
        </p>
        <div className="mt-4">
          <FilePicker
            accept={WORKBOOK_FILES.accept}
            description="Drag an Excel workbook here, or choose one with the button below."
            disabled={runState.status === "running" || selection.reading}
            label="Source workbook"
            multiple={false}
            onFiles={(files) => selection.choose(files, runState.reset)}
          />
        </div>
        {selection.reading ? <ReadingFile testId="reading-file" /> : null}
        {selection.file ? (
          <ChosenFile file={selection.file} testId="chosen-file" />
        ) : null}
        {selection.rejected ? (
          <p className="mt-3 text-sm text-fd-accent-foreground" role="alert">
            {selection.rejected}
          </p>
        ) : null}
      </section>
      <section className={sectionClass}>
        <h2 className="text-xl font-bold tracking-[-0.03em]">2. Run</h2>
        <RunControls
          busyLabel="Unprotecting…"
          disabled={!selection.file || selection.reading}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Rewriting the workbook package in the background…"
          runLabel="Run unprotect"
          state={runState}
        />
      </section>
      <ResultsPanel
        fallbackMediaType={WORKBOOK_FILES.fallbackMediaType}
        state={runState}
      />
    </ToolShell>
  );
}
