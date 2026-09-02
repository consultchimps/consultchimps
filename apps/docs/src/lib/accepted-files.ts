/**
 * The single source of truth for what the in-browser tools accept.
 *
 * Every tool page used to carry its own `accept` string, its own predicate,
 * and its own sentence telling a visitor what to drop. Those three drifted:
 * the byte-level workbook operations accept `.xlsm` and the pages did not, so
 * a macro-enabled workbook was rejected by the picker and by drag-and-drop
 * before any operation could see it. Each kind below therefore derives the
 * `accept` attribute, the predicate, and the wording from one list of
 * extensions and media types, which is what makes a repeat of that drift a
 * compile-time or test-time failure rather than a silent one.
 *
 * The predicate accepts a file whose reported media type is one of the kind's,
 * *or* whose name ends in one of its extensions: browsers disagree about what
 * they report for Office packages, and an unrecognised file arrives with an
 * empty `type`.
 *
 * What a tool then does with a `.xlsm` workbook is the operation's business,
 * not this module's. The split keeps the extension and the macro project; the
 * merge keeps them only when the visitor names the output `.xlsm`; the
 * consolidation always writes a fresh `.xlsx`. See each tool's guide.
 */

/** One family of files a tool page accepts. */
export interface AcceptedFileKind {
  /** The value for a file input's `accept` attribute. */
  readonly accept: string;
  /**
   * The kind named as a singular noun phrase with its article, so it reads
   * inside a sentence: "Drag an Excel .xlsx or .xlsm workbook here".
   */
  readonly description: string;
  /** Lower-case extensions, leading dot, in the order the wording names them. */
  readonly extensions: readonly string[];
  /**
   * The media type a download falls back to when an operation's artifact
   * declares none. Operations that can emit more than one kind of package
   * declare the media type per artifact, so this is only the last resort.
   */
  readonly fallbackMediaType: string;
  /** Every media type a browser may report for these extensions. */
  readonly mediaTypes: readonly string[];
  /** The plural noun phrase, for the pickers that take several files. */
  readonly pluralDescription: string;
  /** Whether a picked or dropped file is one of this kind. */
  readonly accepts: (file: File) => boolean;
  /** The name without its extension, unchanged when it carries none of them. */
  readonly stripExtension: (name: string) => string;
}

interface AcceptedFileKindDefinition {
  readonly description: string;
  readonly extensions: readonly string[];
  readonly mediaTypes: readonly string[];
  readonly pluralDescription: string;
}

function createAcceptedFileKind(
  definition: AcceptedFileKindDefinition,
): AcceptedFileKind {
  const { description, extensions, mediaTypes, pluralDescription } = definition;
  // Built from the same list the wording and the `accept` attribute come from,
  // so an added extension cannot reach one of the three and miss the others.
  const extensionPattern = new RegExp(
    `\\.(?:${extensions.map((extension) => extension.slice(1)).join("|")})$`,
    "iu",
  );
  // A File's `type` is lower-cased by the browser, but the registered media
  // type for a macro-enabled workbook is not all lower case
  // ("...sheet.macroEnabled.12"). Comparing the two verbatim would therefore
  // never match a genuine .xlsm, so both sides are folded here while the
  // canonical spelling is what the page still advertises and downloads with.
  const foldedMediaTypes = mediaTypes.map((mediaType) =>
    mediaType.toLowerCase(),
  );

  return {
    accept: [...mediaTypes, ...extensions].join(","),
    accepts: (file: File): boolean =>
      foldedMediaTypes.includes(file.type.toLowerCase()) ||
      extensionPattern.test(file.name),
    description,
    extensions,
    fallbackMediaType: mediaTypes[0]!,
    mediaTypes,
    pluralDescription,
    stripExtension: (name: string): string =>
      name.replace(extensionPattern, ""),
  };
}

/**
 * Excel workbooks. Both extensions reach the same byte-level operations: the
 * package readers are format-neutral, and the split refuses only a package
 * whose declared type contradicts its name.
 */
export const WORKBOOK_FILES = createAcceptedFileKind({
  description: "an Excel .xlsx or .xlsm workbook",
  extensions: [".xlsx", ".xlsm"],
  mediaTypes: [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel.sheet.macroEnabled.12",
  ],
  pluralDescription: "Excel .xlsx or .xlsm workbooks",
});

/** PowerPoint presentations. The populate engine reads `.pptx` packages only. */
export const PRESENTATION_FILES = createAcceptedFileKind({
  description: "a PowerPoint .pptx presentation",
  extensions: [".pptx"],
  mediaTypes: [
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ],
  pluralDescription: "PowerPoint .pptx presentations",
});

/**
 * Column mapping documents. The mapping is a versioned JSON file rather than a
 * package, and this surface has no filesystem, so the page reads the text and
 * validates the parsed document before any workbook is opened.
 */
export const MAPPING_FILES = createAcceptedFileKind({
  description: "a .json column mapping",
  extensions: [".json"],
  mediaTypes: ["application/json"],
  pluralDescription: ".json column mappings",
});

/** PDF documents. */
export const PDF_FILES = createAcceptedFileKind({
  description: "a .pdf document",
  extensions: [".pdf"],
  mediaTypes: ["application/pdf"],
  pluralDescription: ".pdf documents",
});
