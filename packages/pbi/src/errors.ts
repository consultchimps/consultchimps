import { ConsultChimpsError } from "@consultchimps/core";

/** Refusal codes reserved by ADR 0004, including later decoding/export stages. */
export type PbiErrorCode =
  | "PBI_INVALID_OPTIONS"
  | "PBI_RUNTIME_UNAVAILABLE"
  | "PBI_INVALID_CONTAINER"
  | "PBI_MODEL_UNREADABLE"
  | "PBI_NO_MODEL"
  | "PBI_MODEL_ENCRYPTED"
  | "PBI_NO_EXPORTABLE_TABLES"
  | "PBI_EXPORT_LIMIT_EXCEEDED";

export function invalidContainer(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_INVALID_CONTAINER" satisfies PbiErrorCode,
    "The file is not a supported Power BI ZIP container, or its required parts are damaged or missing. Save a new .pbix in Power BI Desktop and try again.",
    { details: { stage: "container" } },
  );
}

export function unreadableModel(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_MODEL_UNREADABLE" satisfies PbiErrorCode,
    "The embedded model part is empty or damaged. Save a new .pbix with imported data in Power BI Desktop and try again.",
    { details: { stage: "model-part" } },
  );
}

export function encryptedModel(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_MODEL_ENCRYPTED" satisfies PbiErrorCode,
    "The embedded model is encrypted or password-protected. Ask the file owner for an unencrypted .pbix saved with imported data.",
    { details: { stage: "model-part" } },
  );
}
