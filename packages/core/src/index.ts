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

export interface ChimpconsErrorOptions {
  cause?: unknown;
  details?: Record<string, unknown>;
}

export class ChimpconsError extends Error {
  readonly code: string;
  readonly details: Record<string, unknown> | undefined;

  constructor(
    code: string,
    message: string,
    options: ChimpconsErrorOptions = {},
  ) {
    super(message, { cause: options.cause });
    this.name = "ChimpconsError";
    this.code = code;
    this.details = options.details;
  }
}

export function isChimpconsError(error: unknown): error is ChimpconsError {
  return error instanceof ChimpconsError;
}
