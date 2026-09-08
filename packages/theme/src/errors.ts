/**
 * The theme package is runtime-neutral with zero dependencies, so it cannot
 * reach for `@consultchimps/core`'s `ConsultChimpsError`. This local error keeps
 * the same shape a consumer expects: a stable, namespaced `code` beside a
 * human-readable message. It is thrown only for programming mistakes, such as
 * asking for a palette slot that does not exist or passing a malformed colour.
 * A validation pass never throws; it returns a structured report instead.
 */
export class ThemeError extends Error {
  readonly code: string;
  readonly details: Record<string, unknown> | undefined;

  constructor(
    code: string,
    message: string,
    details?: Record<string, unknown>,
  ) {
    super(message);
    this.name = "ThemeError";
    this.code = code;
    this.details = details;
  }
}

export function isThemeError(error: unknown): error is ThemeError {
  return error instanceof ThemeError;
}
