/**
 * The wire contract between the tool pages and the operation Web Worker.
 *
 * Only this module is shared by both sides, and it deliberately contains only
 * types, with no engine imports. The worker resolves a task to
 * the matching byte-level operation and loads that engine on demand, so a page
 * that is merely open never downloads a PDF or workbook engine.
 *
 * Two constraints shape the shapes below:
 *
 * - An `AbortSignal` cannot be structure-cloned, so cancellation travels as a
 *   `cancel` command and the worker aborts its own controller.
 * - `ConsultChimpsError` arrives on the other side as a plain object, so a
 *   failure travels as its message and code and is rebuilt by the client.
 */
import type {
  OperationPlan,
  OperationProgress,
  OperationResult,
} from "@consultchimps/core";
// Type-only imports: the runtime modules are loaded inside the worker.
import type { MergePdfsMetric, SplitPdfMetric } from "@consultchimps/pdf/bytes";
import type {
  PopulatePowerPointTemplateMetric,
  PopulatePowerPointTemplatePlanMetric,
  PresentationInspectionOutcome,
} from "@consultchimps/pptx/bytes";
import type {
  ColumnMapping,
  ColumnMappingSuggestion,
} from "@consultchimps/tabular";
import type {
  ConsolidateWorkbooksMetric,
  CsvReadOptions,
  MergeWorkbooksMetric,
  SplitWorkbookByColumnPlanMetric,
  SplitWorkbookBytesOutcome,
  WorkbookDescriptionOutcome,
  UnprotectWorkbookMetric,
} from "@consultchimps/xlsx/bytes";

/**
 * One input as the visitor chose it. A `File` crosses into the worker as a
 * reference to the same file, not a copy of its bytes, and the worker reads it.
 */
export interface NamedFile {
  readonly name: string;
  readonly file: Blob;
}

/**
 * One output as the page receives it. The worker wraps each output in a
 * `Blob`, which crosses back as a reference the browser manages, so the page
 * never holds the bytes in a buffer of its own.
 */
export interface OutputFile {
  readonly name: string;
  readonly blob: Blob;
  readonly mediaType?: string | undefined;
  /**
   * Set when the browser gave no private file system, so the output is held
   * in memory rather than on disk.
   */
  readonly inMemory?: boolean | undefined;
}

/** A byte-level operation's outcome with its outputs as `OutputFile`s. */
export interface FileOperationOutcome<TMetric extends string> {
  readonly result: OperationResult<TMetric>;
  readonly outputs: readonly OutputFile[];
}

/**
 * The template and records selection a population reads. The two tasks that
 * populate, the preview and the run, take exactly the same options, so the
 * page builds one object and sends it to both.
 */
export interface PresentationPopulateOptions {
  readonly headerRow?: number | undefined;
  readonly outputName?: string | undefined;
  readonly templateSlide?: number | undefined;
  readonly worksheet?: string | undefined;
}

/**
 * The workbook-split selection and formatting options the page exposes.
 *
 * Which engine these reach is decided by the selection fields, not by a mode
 * flag: leaving `table`, `range`, and `sheet` unset while `preserveWorkbook`
 * is not `false` runs the all-worksheet split, where every worksheet carrying
 * the column is filtered in place and the rest of the workbook travels with
 * each output. Naming any one source, or passing `preserveWorkbook: false`,
 * selects a single-source split instead, and that is the only case where
 * `includeBlank` and `includeHiddenSheets` mean anything, because the
 * all-worksheet engine ignores both.
 */
export interface WorkbookSplitOptions {
  readonly column: string;
  /** How to read the input when it is a CSV file (ADR 0007). */
  readonly csv?: CsvReadOptions | undefined;
  readonly filenamePrefix?: string | undefined;
  readonly headerRow?: number | undefined;
  readonly includeBlank?: boolean | undefined;
  readonly includeHiddenSheets?: boolean | undefined;
  readonly preserveWorkbook?: boolean | undefined;
  readonly range?: string | undefined;
  readonly sheet?: string | undefined;
  readonly strict?: boolean | undefined;
  readonly table?: string | undefined;
  readonly values?: boolean | undefined;
}

/**
 * What a workbook inspection is allowed to vary.
 *
 * Both fields mean what they mean to every other workbook reader here, so a
 * page that already collects a header row or a hidden-worksheet choice can
 * hand the inspector the same values it hands the operation it is previewing.
 * The sample bound is deliberately absent: the report renders whatever the
 * library's own limit allows, and a page that asked for fewer samples than the
 * split page shows would make two views of one workbook disagree.
 */
export interface WorkbookInspectOptions {
  /** How to read the file when it is a CSV file. */
  readonly csv?: CsvReadOptions | undefined;
  readonly headerRow?: number | undefined;
  readonly includeHiddenSheets?: boolean | undefined;
}

export type OperationTask =
  | {
      readonly kind: "pdf.plan-split";
      readonly input: NamedFile;
      readonly filenamePrefix?: string | undefined;
    }
  | {
      readonly kind: "pdf.split";
      readonly input: NamedFile;
      readonly filenamePrefix?: string | undefined;
    }
  | {
      readonly kind: "pdf.merge";
      readonly inputs: readonly NamedFile[];
      readonly outputName?: string | undefined;
    }
  | {
      readonly kind: "xlsx.plan-split";
      readonly input: NamedFile;
      readonly options: WorkbookSplitOptions;
    }
  | {
      readonly kind: "xlsx.split";
      readonly input: NamedFile;
      readonly options: WorkbookSplitOptions;
    }
  | {
      readonly kind: "xlsx.merge";
      readonly inputs: readonly NamedFile[];
      // How to read the inputs that are CSV files (ADR 0007).
      readonly csv?: CsvReadOptions | undefined;
      readonly outputName?: string | undefined;
      readonly values?: boolean | undefined;
    }
  | {
      // Rows are stacked in the order the inputs are listed, so the array
      // carries the visitor's arrangement, not an incidental read order.
      readonly kind: "xlsx.consolidate";
      readonly inputs: readonly NamedFile[];
      readonly addSourceColumns?: boolean | undefined;
      // How to read the inputs that are CSV files (ADR 0007).
      readonly csv?: CsvReadOptions | undefined;
      readonly includeHiddenSheets?: boolean | undefined;
      // A parsed and validated version 1 mapping. This surface has no
      // filesystem, so the page reads the document and the operation validates
      // it again before a workbook is opened.
      readonly mapping?: ColumnMapping | undefined;
      readonly normalizeHeaders?: boolean | undefined;
      readonly outputName?: string | undefined;
    }
  | {
      // Draft a mapping from the headers of the tables a consolidation would
      // read, and answer with the draft alone. The operation behind it is the
      // consolidation itself in suggest mode, so the browser's proposal is the
      // library's proposal over the same tables, never a second grouping rule
      // living in the page. Nothing it returns is applied to anything.
      readonly kind: "xlsx.suggest-mapping";
      readonly inputs: readonly NamedFile[];
      readonly csv?: CsvReadOptions | undefined;
      readonly includeHiddenSheets?: boolean | undefined;
    }
  | {
      readonly kind: "xlsx.columns";
      readonly input: NamedFile;
      readonly csv?: CsvReadOptions | undefined;
      readonly headerRow?: number | undefined;
      readonly worksheet?: string | undefined;
    }
  | {
      // Creates nothing: the answer is the description itself, which is why
      // this sits with the plan and column tasks rather than with the byte
      // operations below.
      readonly kind: "xlsx.inspect";
      readonly input: NamedFile;
      readonly options: WorkbookInspectOptions;
    }
  | {
      readonly kind: "xlsx.unprotect";
      readonly input: NamedFile;
      readonly outputName?: string | undefined;
    }
  | {
      readonly kind: "pptx.inspect";
      readonly template: NamedFile;
      readonly templateSlide?: number | undefined;
    }
  | {
      readonly kind: "pptx.plan-populate";
      readonly template: NamedFile;
      readonly workbook: NamedFile;
      readonly options: PresentationPopulateOptions;
    }
  | {
      readonly kind: "pptx.populate";
      readonly template: NamedFile;
      readonly workbook: NamedFile;
      readonly options: PresentationPopulateOptions;
    };

/**
 * The subset of tasks that produce output bytes, which is what the run hook
 * and the results panel deal in. Plan and inspection tasks are excluded.
 */
export type ByteOperationTask = Extract<
  OperationTask,
  {
    kind:
      | "pdf.merge"
      | "pdf.split"
      | "pptx.populate"
      | "xlsx.consolidate"
      | "xlsx.merge"
      | "xlsx.split"
      | "xlsx.unprotect";
  }
>;

/** The column headers a workbook's chosen worksheet offers to a split. */
export interface WorksheetColumns {
  readonly columns: readonly string[];
  readonly worksheet: string;
}

interface OperationTaskResults {
  "pdf.plan-split": OperationPlan<SplitPdfMetric>;
  "pdf.split": FileOperationOutcome<SplitPdfMetric>;
  "pdf.merge": FileOperationOutcome<MergePdfsMetric>;
  "xlsx.plan-split": OperationPlan<SplitWorkbookByColumnPlanMetric>;
  // The split reports more than a plain byte outcome: in all-worksheet mode
  // its result also carries per-output and per-worksheet detail, so the
  // operation's own result type is kept rather than flattened.
  "xlsx.split": {
    readonly result: SplitWorkbookBytesOutcome["result"];
    readonly outputs: readonly OutputFile[];
  };
  "xlsx.merge": FileOperationOutcome<MergeWorkbooksMetric>;
  "xlsx.consolidate": FileOperationOutcome<ConsolidateWorkbooksMetric>;
  // The drafted mapping and the evidence behind each of its entries. The
  // consolidated workbook the suggesting run also builds is dropped in the
  // worker: this task exists to answer a question, not to produce a file.
  "xlsx.suggest-mapping": ColumnMappingSuggestion;
  "xlsx.columns": WorksheetColumns;
  // The description travels beside the operation's structured result, so the
  // page can render both the structure and the counts the library derived.
  "xlsx.inspect": WorkbookDescriptionOutcome;
  "xlsx.unprotect": FileOperationOutcome<UnprotectWorkbookMetric>;
  "pptx.inspect": PresentationInspectionOutcome;
  "pptx.plan-populate": OperationPlan<PopulatePowerPointTemplatePlanMetric>;
  "pptx.populate": FileOperationOutcome<PopulatePowerPointTemplateMetric>;
}

/** What a given task resolves to once the worker has run it. */
export type OperationTaskResult<TTask extends OperationTask> = TTask extends {
  kind: infer TKind;
}
  ? TKind extends keyof OperationTaskResults
    ? OperationTaskResults[TKind]
    : never
  : never;

export interface RunCommand {
  readonly type: "run";
  readonly id: number;
  readonly task: OperationTask;
}

export interface CancelCommand {
  readonly type: "cancel";
  readonly id: number;
}

/** Sent when the page goes away, so the worker deletes the outputs it offered. */
export interface ReleaseCommand {
  readonly type: "release";
}

export type WorkerCommand = CancelCommand | ReleaseCommand | RunCommand;

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
  /** Present only for tasks that produce files; recombined with `value`. */
  readonly artifacts?: readonly OutputFile[] | undefined;
}

export interface FailedEvent {
  readonly type: "failed";
  readonly id: number;
  readonly message: string;
  readonly code?: string | undefined;
}

export type WorkerEvent = DoneEvent | FailedEvent | ProgressEvent;
