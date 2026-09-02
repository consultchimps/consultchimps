/**
 * The worker that performs every byte-level operation the tool pages offer.
 *
 * Splitting a large PDF or rewriting a workbook package is seconds of tight,
 * synchronous work. Running it here keeps the tab responsive: the page only
 * ships inputs in, renders progress events, and receives output buffers.
 *
 * One worker serves every format, and each engine is still pulled in with a
 * dynamic `import()` on the first task that needs it, so opening a tool page
 * downloads no engine until someone asks for a preview or a run.
 */
import {
  isConsultChimpsError,
  type ByteArtifact,
  type ByteOperationOutcome,
  type OperationControlOptions,
} from "@consultchimps/core";

import type {
  OperationTask,
  TransferableArtifact,
  WorkerCommand,
  WorkerEvent,
} from "@/lib/operation-tasks";

/**
 * The worker global, typed locally. Pulling in the `webworker` lib would
 * collide with the DOM lib the rest of the app compiles against, and the two
 * members used here are the whole surface.
 */
const scope = self as unknown as {
  addEventListener(
    type: "message",
    listener: (event: MessageEvent<WorkerCommand>) => void,
  ): void;
  postMessage(message: WorkerEvent, transfer?: Transferable[]): void;
};

/** One controller per in-flight task, so a `cancel` command can reach it. */
const controllers = new Map<number, AbortController>();

interface TaskAnswer {
  readonly value: unknown;
  readonly artifacts?: readonly TransferableArtifact[] | undefined;
  readonly transfer: Transferable[];
}

/**
 * Hand an output's bytes over as a buffer that matches it exactly. A view into
 * a larger pool is copied first, because transferring its buffer would move
 * unrelated bytes and detach whatever else pointed at them.
 */
function toTransferable(artifact: ByteArtifact): TransferableArtifact {
  const { bytes } = artifact;
  const exact =
    bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength;
  return {
    name: artifact.name,
    buffer: (exact ? bytes.buffer : bytes.slice().buffer) as ArrayBuffer,
    ...(artifact.mediaType === undefined
      ? {}
      : { mediaType: artifact.mediaType }),
  };
}

function answerWithOutputs<TMetric extends string>(
  outcome: ByteOperationOutcome<TMetric>,
): TaskAnswer {
  const artifacts = outcome.outputs.map(toTransferable);
  // A duplicate entry in the transfer list is a DataCloneError, so the buffers
  // are deduplicated even though distinct outputs normally own distinct ones.
  const transfer = [...new Set(artifacts.map((artifact) => artifact.buffer))];
  return { value: outcome.result, artifacts, transfer };
}

function answerWithValue(value: unknown): TaskAnswer {
  return { value, transfer: [] };
}

async function perform(
  task: OperationTask,
  controls: Required<OperationControlOptions>,
): Promise<TaskAnswer> {
  switch (task.kind) {
    case "pdf.plan-split": {
      const { planSplitPdfBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithValue(
        await planSplitPdfBytes({
          input: { bytes: task.input.bytes, name: task.input.name },
          filenamePrefix: task.filenamePrefix,
        }),
      );
    }
    case "pdf.split": {
      const { splitPdfBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithOutputs(
        await splitPdfBytes({
          ...controls,
          input: { bytes: task.input.bytes, name: task.input.name },
          filenamePrefix: task.filenamePrefix,
        }),
      );
    }
    case "pdf.merge": {
      const { mergePdfsBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithOutputs(
        await mergePdfsBytes({
          ...controls,
          inputs: task.inputs.map((input) => ({
            bytes: input.bytes,
            name: input.name,
          })),
          outputName: task.outputName,
        }),
      );
    }
    case "xlsx.plan-split": {
      const { planSplitWorkbookBytes } =
        await import("@consultchimps/xlsx/bytes");
      return answerWithValue(
        await planSplitWorkbookBytes({
          ...task.options,
          input: { bytes: task.input.bytes, name: task.input.name },
        }),
      );
    }
    case "xlsx.split": {
      const { splitWorkbookBytes } = await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await splitWorkbookBytes({
          ...controls,
          ...task.options,
          input: { bytes: task.input.bytes, name: task.input.name },
        }),
      );
    }
    case "xlsx.merge": {
      const { mergeWorkbooksBytes } = await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await mergeWorkbooksBytes({
          ...controls,
          inputs: task.inputs.map((input) => ({
            bytes: input.bytes,
            name: input.name,
          })),
          outputName: task.outputName,
          values: task.values,
        }),
      );
    }
    case "xlsx.consolidate": {
      const { consolidateWorkbooksBytes } =
        await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await consolidateWorkbooksBytes({
          ...controls,
          inputs: task.inputs.map((input) => ({
            bytes: input.bytes,
            name: input.name,
          })),
          addSourceColumns: task.addSourceColumns,
          includeHiddenSheets: task.includeHiddenSheets,
          mapping: task.mapping,
          normalizeHeaders: task.normalizeHeaders,
          outputName: task.outputName,
        }),
      );
    }
    case "xlsx.suggest-mapping": {
      const { consolidateWorkbooksBytes } =
        await import("@consultchimps/xlsx/bytes");
      // The suggestion is drafted from the tables the consolidation read, so
      // the page proposes exactly what the library proposes for these
      // workbooks and these options. The consolidated workbook it also builds
      // is not returned: the page asked what the headers look like, not for a
      // file, and reading the tables through the operation is what keeps the
      // browser's draft and the command line's draft the same document.
      const { result } = await consolidateWorkbooksBytes({
        ...controls,
        inputs: task.inputs.map((input) => ({
          bytes: input.bytes,
          name: input.name,
        })),
        includeHiddenSheets: task.includeHiddenSheets,
        suggestMapping: true,
      });
      return answerWithValue(result.suggestion);
    }
    case "xlsx.columns": {
      const { readWorksheetRecordsBytes } =
        await import("@consultchimps/xlsx/bytes");
      const records = await readWorksheetRecordsBytes(
        { bytes: task.input.bytes, name: task.input.name },
        { headerRow: task.headerRow, worksheet: task.worksheet },
      );
      return answerWithValue({
        columns: records.columns,
        worksheet: records.worksheet,
      });
    }
    case "xlsx.inspect": {
      const { describeWorkbookBytes } =
        await import("@consultchimps/xlsx/bytes");
      return answerWithValue(
        await describeWorkbookBytes(
          { bytes: task.input.bytes, name: task.input.name },
          { ...controls, ...task.options },
        ),
      );
    }
    case "xlsx.unprotect": {
      const { unprotectWorkbookBytes } =
        await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await unprotectWorkbookBytes({
          ...controls,
          input: task.input,
          ...(task.outputName === undefined
            ? {}
            : { outputName: task.outputName }),
        }),
      );
    }
    case "pptx.inspect": {
      const { inspectPresentationOutcomeBytes } =
        await import("@consultchimps/pptx/bytes");
      return answerWithValue(
        await inspectPresentationOutcomeBytes(
          { bytes: task.template.bytes, name: task.template.name },
          { ...controls, templateSlide: task.templateSlide },
        ),
      );
    }
    case "pptx.plan-populate": {
      const { planPopulatePresentationBytes } =
        await import("@consultchimps/pptx/bytes");
      return answerWithValue(
        await planPopulatePresentationBytes({
          ...controls,
          ...task.options,
          template: { bytes: task.template.bytes, name: task.template.name },
          workbook: { bytes: task.workbook.bytes, name: task.workbook.name },
        }),
      );
    }
    case "pptx.populate": {
      const { populatePresentationBytes } =
        await import("@consultchimps/pptx/bytes");
      return answerWithOutputs(
        await populatePresentationBytes({
          ...controls,
          ...task.options,
          template: { bytes: task.template.bytes, name: task.template.name },
          workbook: { bytes: task.workbook.bytes, name: task.workbook.name },
        }),
      );
    }
  }
  throw new Error("Unsupported operation task");
}

async function execute(id: number, task: OperationTask): Promise<void> {
  const controller = new AbortController();
  controllers.set(id, controller);
  try {
    const answer = await perform(task, {
      onProgress: (progress) => {
        scope.postMessage({ type: "progress", id, progress });
      },
      signal: controller.signal,
    });
    scope.postMessage(
      {
        type: "done",
        id,
        value: answer.value,
        artifacts: answer.artifacts,
      },
      answer.transfer,
    );
  } catch (error) {
    scope.postMessage({
      type: "failed",
      id,
      message:
        error instanceof Error
          ? error.message
          : "An unexpected problem occurred.",
      code: isConsultChimpsError(error) ? error.code : undefined,
    });
  } finally {
    controllers.delete(id);
  }
}

scope.addEventListener("message", (event) => {
  const command = event.data;
  if (command.type === "cancel") {
    controllers.get(command.id)?.abort();
    return;
  }
  void execute(command.id, command.task);
});
