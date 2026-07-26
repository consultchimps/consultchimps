export type ArtifactKind = "file" | "directory";

export interface Artifact {
  kind: ArtifactKind;
  path: string;
  mediaType?: string;
}

export interface OperationResult {
  operation: string;
  artifacts: Artifact[];
  warnings: string[];
  metrics: Record<string, number>;
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
