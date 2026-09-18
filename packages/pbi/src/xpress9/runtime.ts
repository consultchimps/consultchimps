import { ConsultChimpsError } from "@consultchimps/core";
import createXpress9Module from "../../wasm/xpress9.mjs";
import type { Xpress9Module } from "../../wasm/xpress9.mjs";

/**
 * How this host reaches `xpress9.wasm`. Exactly one source, per ADR 0004
 * Decision 2: a locator the host resolves to a fetchable URL, or the bytes
 * themselves. Omit the whole config in Node and the copy that ships with this
 * package is used.
 */
export interface Xpress9RuntimeConfig {
  readonly locateFile?: (fileName: string) => string;
  readonly wasmBinary?: Uint8Array;
}

/**
 * One XPress9 decoder session. Chunks decompressed through the same decoder
 * share the LZ77 history window, which is what an XPress9 backup stream
 * requires, so a caller decoding one stream holds one decoder and decompresses
 * its chunks in order. A stream that needs an independent window, such as each
 * thread group of a multithreaded backup, needs its own `loadXpress9` call.
 *
 * Two consequences of that shared window, both inherent to the format rather
 * than to this wrapper. Chunks must be fed in stream order and exactly once:
 * feeding the same chunk twice is not a decode of the same data, it is a decode
 * against the wrong history, and it fails. And once a decode has failed, the
 * session's history is no longer in a known state, so the decoder is finished:
 * close it rather than continuing on it.
 */
export interface Xpress9Decoder {
  /**
   * Decompresses one framed chunk into exactly `outputSize` bytes and returns
   * them as an independent copy. A short or failed decode is a damaged model.
   */
  decompress(input: Uint8Array, outputSize: number): Uint8Array;
  /**
   * The module's current linear memory in bytes. The host charges a prediction
   * of this before each frame and compares it with the real figure afterwards,
   * so the accounting is never below what the runtime actually took.
   */
  memoryBytes(): number;
  /** Releases the decoder and its buffers. Safe to call more than once. */
  close(): void;
}

/** The single asset name a locator is ever asked for. */
const ASSET_NAME = "xpress9.wasm";

/**
 * Where the committed artifact sits relative to this module, in the two
 * layouts that exist. The first is the built package, where tsup bundles this
 * file into `dist/index.js` and `wasm/` is its sibling. The second is the
 * source tree, where this file is `src/xpress9/runtime.ts`. Both are tried in
 * order and the first that reads wins, so the default path works from the test
 * suite and from an installed package without either guessing at the other.
 */
const DEFAULT_WASM_PATHS = ["../wasm/xpress9.wasm", "../../wasm/xpress9.wasm"];

type LoadStage = "load" | "compile" | "instantiate" | "closed";

/**
 * The package compiles against `lib: ["ES2024"]` and `types: ["node"]`, neither
 * of which declares the WebAssembly namespace, and it is not worth pulling the
 * whole DOM library in for one function. This declares exactly the one call
 * used below, which every host that can run this module provides.
 */
declare const WebAssembly: { validate(bytes: Uint8Array): boolean };

function invalidRuntimeOptions(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_INVALID_OPTIONS",
    "The XPress9 runtime option is not usable. Supply exactly one of a locator function or a nonempty Uint8Array of the xpress9.wasm bytes, or omit the option entirely to use the copy that ships with this package.",
    {
      details: {
        stage: "options",
        invalidOptions: [
          {
            path: "runtime.xpress9",
            requirement: "exactly one locator function or nonempty Uint8Array",
          },
        ],
      },
    },
  );
}

/**
 * Never carries a cause, a URL, or any text from the toolchain or the host, so
 * a locator's own error message cannot reach a log or a user through this path.
 */
function runtimeUnavailable(stage: LoadStage): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_RUNTIME_UNAVAILABLE",
    stage === "closed"
      ? "This XPress9 decompressor has already been closed. Load a new one before decompressing again."
      : "The XPress9 decompressor could not be started. In Node, reinstall or rebuild @consultchimps/pbi so that its xpress9.wasm is present; in a browser, point the xpress9 runtime option at a reachable copy of xpress9.wasm.",
    { details: { runtime: "xpress9", stage } },
  );
}

/** The decoder's own error text is deliberately never read into this. */
function damagedModel(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_MODEL_UNREADABLE",
    "The embedded model is damaged: a compressed block did not decompress to the size it declares. Save a new .pbix with imported data in Power BI Desktop and try again.",
    { details: { stage: "xpress9" } },
  );
}

function memoryExhausted(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_EXPORT_LIMIT_EXCEEDED",
    "The XPress9 decompressor could not grow its memory far enough to decode this model. Use a smaller model, or run the export in a host that can give WebAssembly more memory.",
    { details: { stage: "xpress9", reason: "wasm-memory-exhausted" } },
  );
}

/**
 * Validates before any locator call, any read, and any allocation, so a
 * malformed option never reaches the host.
 */
function selectSource(
  config: Xpress9RuntimeConfig,
):
  | { kind: "bytes"; bytes: Uint8Array }
  | { kind: "locator"; locateFile: (fileName: string) => string } {
  const { locateFile, wasmBinary } = config;
  const hasLocator = locateFile !== undefined;
  const hasBytes = wasmBinary !== undefined;
  if (hasLocator === hasBytes) throw invalidRuntimeOptions();
  if (hasLocator) {
    if (typeof locateFile !== "function") throw invalidRuntimeOptions();
    return { kind: "locator", locateFile };
  }
  if (!(wasmBinary instanceof Uint8Array) || wasmBinary.byteLength === 0) {
    throw invalidRuntimeOptions();
  }
  return { kind: "bytes", bytes: wasmBinary };
}

/**
 * Assembled rather than written as a literal, deliberately. tsup rewrites a
 * literal `node:fs/promises` to a bare `fs/promises`, which a browser bundler
 * would then try to resolve out of `node_modules` and fail the build on. Built
 * at runtime, the specifier is opaque to every bundler: Node resolves it, and a
 * browser leaves a runtime import that the caller below catches and turns into
 * a refusal telling the host to configure the runtime.
 */
const nodeFilesystemModule = ["node", "fs/promises"].join(":");

/**
 * Reads the committed artifact through a dynamic import, so no static Node
 * builtin import exists anywhere in a browser bundle of this package.
 */
async function readDefaultWasm(): Promise<Uint8Array> {
  let readFile: (path: URL) => Promise<Uint8Array>;
  try {
    ({ readFile } = (await import(nodeFilesystemModule)) as unknown as {
      readFile: (path: URL) => Promise<Uint8Array>;
    });
  } catch {
    // No filesystem, so this is a browser or a worker and the caller must
    // configure the runtime. There is no default to fall back to.
    throw runtimeUnavailable("load");
  }
  for (const candidate of DEFAULT_WASM_PATHS) {
    try {
      const bytes = await readFile(new URL(candidate, import.meta.url));
      if (bytes.byteLength > 0) return bytes;
    } catch {
      // Try the next layout.
    }
  }
  throw runtimeUnavailable("load");
}

async function fetchWasm(url: string): Promise<Uint8Array> {
  let response: Response;
  try {
    response = await fetch(url);
  } catch {
    throw runtimeUnavailable("load");
  }
  if (!response.ok) throw runtimeUnavailable("load");
  try {
    return new Uint8Array(await response.arrayBuffer());
  } catch {
    throw runtimeUnavailable("load");
  }
}

async function resolveWasmBytes(
  config: Xpress9RuntimeConfig | undefined,
): Promise<Uint8Array> {
  if (config === undefined) return readDefaultWasm();
  const source = selectSource(config);
  if (source.kind === "bytes") return source.bytes;
  let url: string;
  try {
    url = source.locateFile(ASSET_NAME);
  } catch {
    throw runtimeUnavailable("load");
  }
  if (typeof url !== "string" || url.length === 0) {
    throw runtimeUnavailable("load");
  }
  return fetchWasm(url);
}

/**
 * Loads a fresh XPress9 decoder. Every call instantiates its own WebAssembly
 * module, so two decoders share no linear memory and no history window.
 */
export async function loadXpress9(
  config?: Xpress9RuntimeConfig,
): Promise<Xpress9Decoder> {
  const bytes = await resolveWasmBytes(config);

  // Validation is the whole of what compiling would check, and separating it
  // from instantiation is what makes the compile and instantiate stages
  // distinguishable to the caller.
  if (!WebAssembly.validate(bytes)) throw runtimeUnavailable("compile");

  let module: Xpress9Module;
  try {
    module = await createXpress9Module({ wasmBinary: bytes });
  } catch {
    throw runtimeUnavailable("instantiate");
  }

  const context = module._x9_create();
  if (context === 0) throw runtimeUnavailable("instantiate");

  let closed = false;
  let sourcePointer = 0;
  let sourceCapacity = 0;
  let destinationPointer = 0;
  let destinationCapacity = 0;

  /**
   * Grows one cached buffer in place of allocating one per chunk. A backup
   * stream is thousands of chunks of similar size, so reusing the pair keeps
   * the module's linear memory flat across the stream.
   *
   * Ownership is released before the new allocation is attempted, so a malloc
   * that returns zero leaves no pointer behind. Keeping the old address would
   * hand `close()` an already-freed pointer to free a second time, which traps
   * or corrupts the allocator and buries the capacity refusal underneath it.
   */
  const grow = (which: "source" | "destination", needed: number): void => {
    const capacity = which === "source" ? sourceCapacity : destinationCapacity;
    if (needed <= capacity) return;
    const current = which === "source" ? sourcePointer : destinationPointer;
    if (current !== 0) module._x9_free(current);
    if (which === "source") {
      sourcePointer = 0;
      sourceCapacity = 0;
    } else {
      destinationPointer = 0;
      destinationCapacity = 0;
    }
    const pointer = module._x9_malloc(needed);
    if (pointer === 0) throw memoryExhausted();
    if (which === "source") {
      sourcePointer = pointer;
      sourceCapacity = needed;
    } else {
      destinationPointer = pointer;
      destinationCapacity = needed;
    }
  };

  const release = (): void => {
    if (sourcePointer !== 0) module._x9_free(sourcePointer);
    if (destinationPointer !== 0) module._x9_free(destinationPointer);
    sourcePointer = 0;
    sourceCapacity = 0;
    destinationPointer = 0;
    destinationCapacity = 0;
  };

  return {
    decompress(input: Uint8Array, outputSize: number): Uint8Array {
      if (closed) throw runtimeUnavailable("closed");
      // A declared size the framing could not have produced is damaged framing,
      // not a caller mistake, so it reads as a damaged model.
      if (!Number.isSafeInteger(outputSize) || outputSize <= 0) {
        throw damagedModel();
      }
      if (input.byteLength === 0) throw damagedModel();

      grow("source", input.byteLength);
      grow("destination", outputSize);

      // Read HEAPU8 fresh after every allocation: ALLOW_MEMORY_GROWTH detaches
      // and replaces the view whenever linear memory grows.
      module.HEAPU8.set(input, sourcePointer);
      const written = module._x9_decompress(
        context,
        sourcePointer,
        input.byteLength,
        destinationPointer,
        outputSize,
      );
      if (written !== outputSize) throw damagedModel();
      return module.HEAPU8.slice(
        destinationPointer,
        destinationPointer + outputSize,
      );
    },
    memoryBytes(): number {
      return module.HEAPU8.byteLength;
    },
    close(): void {
      if (closed) return;
      closed = true;
      release();
      module._x9_destroy(context);
    },
  };
}
