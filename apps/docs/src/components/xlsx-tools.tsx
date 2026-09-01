"use client";
import { useCallback, useState } from "react";
import { Ban, Download, FileText, LoaderCircle } from "lucide-react";
import {
  useOperationRun,
  primaryButtonClass,
  secondaryButtonClass,
  ToolShell,
  sectionClass,
} from "./tool-kit";

const isExcel = (name: string) => /\.(xlsx|xlsm)$/iu.test(name.trim());
export function XlsxUnprotectTool() {
  const [file, setFile] = useState<{ name: string; bytes: Uint8Array } | null>(
    null,
  );
  const runState = useOperationRun();
  const choose = (candidate: File | undefined) => {
    if (!candidate || !isExcel(candidate.name)) {
      setFile(null);
      runState.reset();
      return;
    }
    void candidate
      .arrayBuffer()
      .then((buffer) =>
        setFile({ name: candidate.name.trim(), bytes: new Uint8Array(buffer) }),
      );
  };
  const start = useCallback(() => {
    if (!file) return;
    void runState.run({
      kind: "xlsx.unprotect",
      input: file,
      outputName: file.name.replace(/\.(xlsx|xlsm)$/iu, "-unprotected.$1"),
    });
  }, [file, runState]);
  return (
    <ToolShell
      kicker="Online tool · Excel unprotect"
      title="Unprotect an Excel workbook"
      description="Remove ordinary worksheet and workbook-structure protection locally. Nothing is uploaded and the original workbook is untouched."
      guideHref="/docs/tools/excel-unprotect"
      guideLabel="Read the unprotect guide"
    >
      <section className={sectionClass}>
        <h2 className="text-xl font-bold">Choose a workbook</h2>
        <p className="mt-2 text-sm text-fd-muted-foreground">
          Supported formats: .xlsx and .xlsm. Encrypted files that require a
          password to open are not supported.
        </p>
        <label
          className="mt-4 inline-flex cursor-pointer items-center rounded-lg border bg-fd-card px-5 py-3 text-sm font-semibold hover:bg-fd-accent"
          htmlFor="xlsx-unprotect-input"
        >
          Choose Excel file
          <input
            className="sr-only"
            id="xlsx-unprotect-input"
            accept=".xlsx,.xlsm"
            type="file"
            onChange={(event) => choose(event.target.files?.[0])}
          />
        </label>
        {file ? (
          <p className="mt-4 flex items-center gap-2 text-sm">
            <FileText className="size-4" />
            {file.name}
          </p>
        ) : null}
      </section>
      <section className={sectionClass}>
        <h2 className="text-xl font-bold">Run</h2>
        <div className="mt-4 flex gap-3">
          <button
            className={primaryButtonClass}
            disabled={!file || runState.status === "running"}
            onClick={start}
            type="button"
          >
            {runState.status === "running" ? (
              <LoaderCircle className="size-4 animate-spin" />
            ) : (
              <Download className="size-4" />
            )}
            Run unprotect
          </button>
          <button
            className={secondaryButtonClass}
            disabled={runState.status !== "running"}
            onClick={runState.cancel}
            type="button"
          >
            <Ban className="size-4" />
            Cancel
          </button>
        </div>
        {runState.progress ? (
          <p className="mt-4 text-sm">{runState.progress.stage}…</p>
        ) : null}
        <pre className="mt-4 whitespace-pre-wrap text-sm">
          {runState.message}
        </pre>
        {runState.outputs.map((output) => (
          <button
            className={`${secondaryButtonClass} mt-3`}
            key={output.name}
            onClick={() => {
              const copy = new Uint8Array(output.bytes);
              const url = URL.createObjectURL(new Blob([copy.buffer]));
              const link = document.createElement("a");
              link.href = url;
              link.download = output.name;
              link.click();
              URL.revokeObjectURL(url);
            }}
            type="button"
          >
            Download {output.name}
          </button>
        ))}
      </section>
    </ToolShell>
  );
}
