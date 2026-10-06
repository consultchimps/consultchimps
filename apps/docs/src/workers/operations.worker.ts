/**
 * The worker that performs every byte-level operation the tool pages offer.
 *
 * Splitting a large PDF or rewriting a workbook package is seconds of tight,
 * synchronous work. Running it here keeps the tab responsive: the page only
 * hands over the chosen `File`s, renders progress events, and receives the
 * outputs as `Blob`s. Reading each input happens here, never in the page.
 *
 * One worker serves every format, and each engine is still pulled in with a
 * dynamic `import()` on the first task that needs it, so opening a tool page
 * downloads no engine until someone asks for a preview or a run.
 */
import {
  ConsultChimpsError,
  isConsultChimpsError,
  OPERATION_ABORTED,
  type ByteArtifact,
  type ByteOperationOutcome,
  type OperationControlOptions,
} from "@consultchimps/core";

import type {
  NamedFile,
  OperationTask,
  OutputFile,
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
  postMessage(message: WorkerEvent): void;
};

/** One controller per in-flight task, so a `cancel` command can reach it. */
const controllers = new Map<number, AbortController>();

interface TaskAnswer {
  readonly value: unknown;
  readonly artifacts?: readonly OutputFile[] | undefined;
}

/** The code a chosen file that can no longer be read fails with. */
const FILE_UNREADABLE = "FILE_UNREADABLE";

/** Said when a chosen file can no longer be read, as the pickers once said it. */
const UNREADABLE_FILE =
  "could not be read. It may have moved, gone offline, or been removed since it was chosen. Choose it again, or pick another file";

/** A promise that rejects with the usual cancellation once `signal` aborts. */
function cancelled(signal: AbortSignal | undefined): Promise<never> {
  return new Promise((_, reject) => {
    if (signal === undefined) return;
    const fail = (): void => {
      reject(
        new ConsultChimpsError(
          OPERATION_ABORTED,
          "The task was cancelled before it finished. No source data was changed and no partial output was produced.",
        ),
      );
    };
    if (signal.aborted) fail();
    else signal.addEventListener("abort", fail, { once: true });
  });
}

/**
 * Read one input whole, for the operations that need every byte at once. The
 * copy lives here in the worker only; the page never held one.
 */
async function whole(
  input: NamedFile,
  signal: AbortSignal | undefined,
): Promise<{
  readonly name: string;
  readonly bytes: Uint8Array;
}> {
  // A large or cloud-backed file can take a while; Cancel must not wait for it.
  const reading = input.file.arrayBuffer().then(
    (buffer) => ({ name: input.name, bytes: new Uint8Array(buffer) }),
    (error: unknown) => {
      throw new ConsultChimpsError(
        FILE_UNREADABLE,
        `"${input.name}" ${UNREADABLE_FILE}`,
        { cause: error, details: { source: input.name } },
      );
    },
  );
  return Promise.race([reading, cancelled(signal)]);
}

/**
 * Hand an output over as a `Blob`. The page receives a reference to it, and
 * the buffer here is dropped once the task answers.
 */
function toOutputFile(artifact: ByteArtifact): OutputFile {
  return {
    name: artifact.name,
    blob: new Blob([artifact.bytes as Uint8Array<ArrayBuffer>], {
      type: artifact.mediaType ?? "",
    }),
    ...(artifact.mediaType === undefined
      ? {}
      : { mediaType: artifact.mediaType }),
  };
}

function answerWithOutputs<TMetric extends string>(
  outcome: ByteOperationOutcome<TMetric>,
): TaskAnswer {
  return {
    value: outcome.result,
    artifacts: outcome.outputs.map(toOutputFile),
  };
}

function answerWithValue(value: unknown): TaskAnswer {
  return { value };
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
          input: await whole(task.input, controls.signal),
          filenamePrefix: task.filenamePrefix,
        }),
      );
    }
    case "pdf.split": {
      const { splitPdfBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithOutputs(
        await splitPdfBytes({
          ...controls,
          input: await whole(task.input, controls.signal),
          filenamePrefix: task.filenamePrefix,
        }),
      );
    }
    case "pdf.merge": {
      const { mergePdfsBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithOutputs(
        await mergePdfsBytes({
          ...controls,
          inputs: await Promise.all(
            task.inputs.map((input) => whole(input, controls.signal)),
          ),
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
          input: await whole(task.input, controls.signal),
        }),
      );
    }
    case "xlsx.split": {
      const { splitWorkbookBytes } = await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await splitWorkbookBytes({
          ...controls,
          ...task.options,
          input: await whole(task.input, controls.signal),
        }),
      );
    }
    case "xlsx.merge": {
      const { mergeWorkbooksBytes } = await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await mergeWorkbooksBytes({
          ...controls,
          inputs: await Promise.all(
            task.inputs.map((input) => whole(input, controls.signal)),
          ),
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
          inputs: await Promise.all(
            task.inputs.map((input) => whole(input, controls.signal)),
          ),
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
        inputs: await Promise.all(
          task.inputs.map((input) => whole(input, controls.signal)),
        ),
        includeHiddenSheets: task.includeHiddenSheets,
        suggestMapping: true,
      });
      return answerWithValue(result.suggestion);
    }
    case "xlsx.columns": {
      const { readWorksheetRecordsBytes } =
        await import("@consultchimps/xlsx/bytes");
      const records = await readWorksheetRecordsBytes(
        await whole(task.input, controls.signal),
        {
          headerRow: task.headerRow,
          worksheet: task.worksheet,
        },
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
        await describeWorkbookBytes(await whole(task.input, controls.signal), {
          ...controls,
          ...task.options,
        }),
      );
    }
    case "xlsx.unprotect": {
      const { unprotectWorkbookBytes } =
        await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await unprotectWorkbookBytes({
          ...controls,
          input: await whole(task.input, controls.signal),
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
          await whole(task.template, controls.signal),
          {
            ...controls,
            templateSlide: task.templateSlide,
          },
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
          template: await whole(task.template, controls.signal),
          workbook: await whole(task.workbook, controls.signal),
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
          template: await whole(task.template, controls.signal),
          workbook: await whole(task.workbook, controls.signal),
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
    scope.postMessage({
      type: "done",
      id,
      value: answer.value,
      artifacts: answer.artifacts,
    });
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
