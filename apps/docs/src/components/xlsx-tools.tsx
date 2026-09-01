"use client";
import { useState } from "react";
import { Download } from "lucide-react";

export function XlsxUnprotectTool() {
  const [file, setFile] = useState<File | null>(null);
  const [message, setMessage] = useState("");
  const run = async () => {
    if (!file) return;
    try {
      const { unprotectWorkbookBytes } =
        await import("@consultchimps/xlsx/bytes");
      const outcome = await unprotectWorkbookBytes({
        input: {
          name: file.name,
          bytes: new Uint8Array(await file.arrayBuffer()),
        },
        outputName: file.name.replace(/\.(xlsx|xlsm)$/iu, "-unprotected.$1"),
      });
      const artifact = outcome.outputs[0]!;
      const copy = new Uint8Array(artifact.bytes.byteLength);
      copy.set(artifact.bytes);
      const url = URL.createObjectURL(
        new Blob([copy.buffer], {
          type:
            artifact.mediaType ??
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
      );
      const link = document.createElement("a");
      link.href = url;
      link.download = artifact.name;
      link.click();
      URL.revokeObjectURL(url);
      setMessage(
        `Removed ${outcome.result.metrics.sheetProtectionsRemoved} worksheet and ${outcome.result.metrics.workbookProtectionsRemoved} workbook protection element(s). Your original file was not changed.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Could not process this workbook.",
      );
    }
  };
  return (
    <main className="mx-auto w-full max-w-[900px] flex-1 px-6 py-14">
      <p className="manual-kicker">Online tool · Excel unprotect</p>
      <h1 className="mt-6 text-4xl font-bold">Unprotect an Excel workbook</h1>
      <p className="mt-5 text-lg leading-8 text-fd-muted-foreground">
        Everything runs locally in your browser. Nothing is uploaded, and your
        original file is untouched.
      </p>
      <section className="mt-10 rounded-xl border p-6">
        <label className="block font-semibold" htmlFor="xlsx-input">
          Choose an .xlsx or .xlsm workbook
        </label>
        <input
          className="mt-4"
          id="xlsx-input"
          accept=".xlsx,.xlsm"
          type="file"
          onChange={(event) => setFile(event.target.files?.[0] ?? null)}
        />
        <button
          className="mt-6 inline-flex items-center gap-2 rounded-lg bg-fd-primary px-4 py-2 text-fd-primary-foreground disabled:opacity-50"
          disabled={!file}
          onClick={() => void run()}
          type="button"
        >
          <Download className="size-4" />
          Run unprotect
        </button>
        {message ? (
          <p className="mt-5 whitespace-pre-wrap text-sm">{message}</p>
        ) : null}
      </section>
    </main>
  );
}
