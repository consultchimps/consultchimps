export type ArtifactKind = "file" | "directory";

export interface Artifact {
  kind: ArtifactKind;
  path: string;
  mediaType?: string;
}

/**
 * The structured outcome of a completed operation. TMetric narrows the metric
 * names an operation reports, making metric renames a compile-time error; the
 * default keeps consumers that handle arbitrary operations working.
 */
export interface OperationResult<TMetric extends string = string> {
  operation: string;
  artifacts: Artifact[];
  warnings: string[];
  metrics: Record<TMetric, number>;
}

export interface OperationProgress {
  operation: string;
  stage: string;
  completed: number;
  total: number;
  detail?: string;
}

export type ProgressReporter = (progress: OperationProgress) => void;

/**
 * Cross-cutting controls accepted by every long-running operation. Progress
 * events are deterministic for identical inputs and options; aborting stops
 * the operation before its next unit of work with an OPERATION_ABORTED error.
 */
export interface OperationControlOptions {
  signal?: AbortSignal | undefined;
  onProgress?: ProgressReporter | undefined;
}

export interface PlannedOutput {
  kind: ArtifactKind;
  path: string;
  /**
   * True when something already exists at this path. Executing the operation
   * without overwrite enabled will refuse to replace it.
   */
  exists: boolean;
  mediaType?: string;
}

/**
 * The validated write plan for an operation: every input it will read and
 * every output it intends to create, computed without writing anything.
 */
export interface OperationPlan<TMetric extends string = string> {
  operation: string;
  inputs: string[];
  outputs: PlannedOutput[];
  warnings: string[];
  metrics: Record<TMetric, number>;
}

export const OPERATION_ABORTED = "OPERATION_ABORTED";

export function throwIfAborted(
  signal: AbortSignal | undefined,
  operation: string,
): void {
  if (signal?.aborted) {
    throw new ConsultChimpsError(
      OPERATION_ABORTED,
      `The "${operation}" operation was cancelled before it finished. No source file was changed; output files completed before the cancellation may remain.`,
      {
        cause: signal.reason,
        details: { operation },
      },
    );
  }
}

export interface ConsultChimpsErrorOptions {
  cause?: unknown;
  details?: Record<string, unknown>;
}

export class ConsultChimpsError extends Error {
  readonly code: string;
  readonly details: Record<string, unknown> | undefined;

  constructor(
    code: string,
    message: string,
    options: ConsultChimpsErrorOptions = {},
  ) {
    super(message, { cause: options.cause });
    this.name = "ConsultChimpsError";
    this.code = code;
    this.details = options.details;
  }
}

export function isConsultChimpsError(
  error: unknown,
): error is ConsultChimpsError {
  return error instanceof ConsultChimpsError;
}
