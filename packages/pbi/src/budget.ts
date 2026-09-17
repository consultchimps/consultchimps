import { ConsultChimpsError } from "@consultchimps/core";

/** Limits for the container reader, not the future browser export pipeline. */
export interface PbiContainerOptions {
  inputBytes?: number;
  decodedBytes?: number;
  peakBytes?: number;
}

export interface ContainerLimits {
  inputBytes: number;
  decodedBytes: number;
  peakBytes: number;
}

// Conservative reader ceilings. Browser export defaults need end-to-end corpus
// measurements, including the runtimes and workbook writer, before shipping.
const defaults: ContainerLimits = {
  inputBytes: 64 * 1024 * 1024,
  decodedBytes: 256 * 1024 * 1024,
  peakBytes: 768 * 1024 * 1024,
};

export function validateLimits(options: PbiContainerOptions): ContainerLimits {
  const paths = ["inputBytes", "decodedBytes", "peakBytes"] as const;
  const invalidOptions = paths
    .filter((path) => {
      const value = options[path];
      return (
        value !== undefined && (!Number.isSafeInteger(value) || value <= 0)
      );
    })
    .map((path) => ({
      path,
      requirement: "a positive safe-integer byte count",
    }));
  if (invalidOptions.length > 0) {
    throw new ConsultChimpsError(
      "PBI_INVALID_OPTIONS",
      "One or more byte limits are invalid. Omit each option or supply a positive safe-integer byte count.",
      { details: { stage: "options", invalidOptions } },
    );
  }
  return {
    inputBytes: options.inputBytes ?? defaults.inputBytes,
    decodedBytes: options.decodedBytes ?? defaults.decodedBytes,
    peakBytes: options.peakBytes ?? defaults.peakBytes,
  };
}

export function checkLimit(
  limits: ContainerLimits,
  option: keyof ContainerLimits,
  required: number,
  stage: "container" | "model-part",
): void {
  if (!Number.isSafeInteger(required) || required > limits[option]) {
    throw new ConsultChimpsError(
      "PBI_EXPORT_LIMIT_EXCEEDED",
      "Reading the Power BI container would exceed a configured byte limit. Use a smaller model, or explicitly increase the limit if your environment can support it.",
      { details: { stage, option, limit: limits[option], required } },
    );
  }
}

export function unboundedContainer(): ConsultChimpsError {
  // Refused as a limit because inflation or ZIP64 sizes cannot be bounded
  // before allocation, not because any configured byte limit was exceeded.
  return new ConsultChimpsError(
    "PBI_EXPORT_LIMIT_EXCEEDED",
    "This Power BI file uses a ZIP layout the reader cannot bound: the model part must be stored without ZIP compression in a single-volume, non-ZIP64 archive. Save the .pbix again in Power BI Desktop and try again.",
    { details: { stage: "container", reason: "unsupported-zip-layout" } },
  );
}
