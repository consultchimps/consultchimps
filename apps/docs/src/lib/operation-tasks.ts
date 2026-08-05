/**
 * The wire contract between the tool pages and the operation Web Worker.
 *
 * Only this module is shared by both sides, and it deliberately contains types
 * plus one tiny helper — no engine imports. The worker resolves a task to the
 * matching byte-level operation and loads that engine on demand, so a page that
 * is merely open never downloads a PDF or workbook engine.
 *
 * Two constraints shape the shapes below:
 *
 * - An `AbortSignal` cannot be structure-cloned, so cancellation travels as a
 *   `cancel` command and the worker aborts its own controller.
 * - `ConsultChimpsError` arrives on the other side as a plain object, so a
 *   failure travels as its message and code and is rebuilt by the client.
 */
import type {
  ByteArtifact,
  ByteOperationOutcome,
  OperationPlan,
  OperationProgress,
} from "@consultchimps/core";
// Type-only imports: the runtime modules are loaded inside the worker.
import type { MergePdfsMetric, SplitPdfMetric } from "@consultchimps/pdf/bytes";
import type {
  MergeWorkbooksMetric,
  SplitWorkbookByColumnMetric,
  SplitWorkbookByColumnPlanMetric,
} from "@consultchimps/xlsx/bytes";

/** One in-memory input, in the shape every byte-level operation accepts. */
export interface NamedBytes {
  readonly name: string;
  readonly bytes: Uint8Array;
}

/** The workbook-split selection and formatting options the page exposes. */
export interface WorkbookSplitOptions {
  readonly column: string;
  readonly filenamePrefix?: string | undefined;
  readonly headerRow?: number | undefined;
  readonly includeBlank?: boolean | undefined;
  readonly includeHiddenSheets?: boolean | undefined;
  readonly preserveWorkbook?: boolean | undefined;
  readonly range?: string | undefined;
  readonly sheet?: string | undefined;
  readonly table?: string | undefined;
  readonly values?: boolean | undefined;
}

export type OperationTask =
  | {
      readonly kind: "pdf.plan-split";
      readonly input: NamedBytes;
      readonly filenamePrefix?: string | undefined;
    }
  | {
      readonly kind: "pdf.split";
      readonly input: NamedBytes;
      readonly filenamePrefix?: string | undefined;
    }
  | {
      readonly kind: "pdf.merge";
      readonly inputs: readonly NamedBytes[];
      readonly outputName?: string | undefined;
    }
  | {
      readonly kind: "xlsx.plan-split";
      readonly input: NamedBytes;
      readonly options: WorkbookSplitOptions;
    }
  | {
      readonly kind: "xlsx.split";
      readonly input: NamedBytes;
      readonly options: WorkbookSplitOptions;
    }
  | {
      readonly kind: "xlsx.merge";
      readonly inputs: readonly NamedBytes[];
      readonly outputName?: string | undefined;
      readonly values?: boolean | undefined;
    }
  | {
      readonly kind: "xlsx.columns";
      readonly input: NamedBytes;
      readonly headerRow?: number | undefined;
      readonly worksheet?: string | undefined;
    };

/**
 * The subset of tasks that produce output bytes, which is what the run hook
 * and the results panel deal in. Plan and inspection tasks are excluded.
 */
export type ByteOperationTask = Extract<
  OperationTask,
  { kind: "pdf.merge" | "pdf.split" | "xlsx.merge" | "xlsx.split" }
>;

/** The column headers a workbook's chosen worksheet offers to a split. */
export interface WorksheetColumns {
  readonly columns: readonly string[];
  readonly worksheet: string;
}

interface OperationTaskResults {
  "pdf.plan-split": OperationPlan<SplitPdfMetric>;
  "pdf.split": ByteOperationOutcome<SplitPdfMetric>;
  "pdf.merge": ByteOperationOutcome<MergePdfsMetric>;
  "xlsx.plan-split": OperationPlan<SplitWorkbookByColumnPlanMetric>;
  "xlsx.split": ByteOperationOutcome<SplitWorkbookByColumnMetric>;
  "xlsx.merge": ByteOperationOutcome<MergeWorkbooksMetric>;
  "xlsx.columns": WorksheetColumns;
}

/** What a given task resolves to once the worker has run it. */
export type OperationTaskResult<TTask extends OperationTask> =
  OperationTaskResults[TTask["kind"]];

export interface RunCommand {
  readonly type: "run";
  readonly id: number;
  readonly task: OperationTask;
}

export interface CancelCommand {
  readonly type: "cancel";
  readonly id: number;
}

export type WorkerCommand = CancelCommand | RunCommand;

/**
 * An output on its way back to the page. The bytes travel as a standalone
 * ArrayBuffer so they can be transferred instead of copied; the buffer always
 * matches the artifact exactly, never a larger pool the view sat inside.
 */
export interface TransferableArtifact {
  readonly name: string;
  readonly buffer: ArrayBuffer;
  readonly mediaType?: string | undefined;
}

export interface ProgressEvent {
  readonly type: "progress";
  readonly id: number;
  readonly progress: OperationProgress;
}

export interface DoneEvent {
  readonly type: "done";
  readonly id: number;
  /**
   * The structure-cloneable part of the answer: a plan, a column list, or the
   * `result` half of a byte-operation outcome.
   */
  readonly value: unknown;
  /** Present only for tasks that produce bytes; recombined with `value`. */
  readonly artifacts?: readonly TransferableArtifact[] | undefined;
}

export interface FailedEvent {
  readonly type: "failed";
  readonly id: number;
  readonly message: string;
  readonly code?: string | undefined;
}

export type WorkerEvent = DoneEvent | FailedEvent | ProgressEvent;

/** Rebuild a page-side artifact from the bytes the worker handed over. */
export function fromTransferable(artifact: TransferableArtifact): ByteArtifact {
  return {
    name: artifact.name,
    bytes: new Uint8Array(artifact.buffer),
    ...(artifact.mediaType === undefined
      ? {}
      : { mediaType: artifact.mediaType }),
  };
}
