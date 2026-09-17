/**
 * Hand-written types for the Emscripten loader that
 * `packages/pbi/scripts/build-wasm.mjs` generates as `xpress9.mjs`.
 *
 * The build does not emit TypeScript declarations, and the generated glue is
 * minified, so this file is the contract between the committed artifact and
 * `packages/pbi/src/xpress9/runtime.ts`. It describes exactly the surface the
 * build script pins through EXPORTED_FUNCTIONS, EXPORTED_RUNTIME_METHODS and
 * INCOMING_MODULE_JS_API, and nothing else. If the build script's export lists
 * change, this file changes with them.
 *
 * Every pointer is a wasm32 byte offset into `HEAPU8`, so a plain number.
 */

/** The two module options INCOMING_MODULE_JS_API allows. */
export interface Xpress9ModuleOptions {
  /** Maps an asset name, always "xpress9.wasm", to a URL the host can fetch. */
  readonly locateFile?: (fileName: string) => string;
  /** The wasm bytes, supplied directly so nothing is fetched. */
  readonly wasmBinary?: Uint8Array;
}

export interface Xpress9Module {
  /**
   * The module's linear memory. ALLOW_MEMORY_GROWTH replaces this view when
   * memory grows, so read it fresh after every allocation rather than caching
   * it across calls.
   */
  readonly HEAPU8: Uint8Array;
  /** Reads a NUL-terminated C string out of linear memory. */
  UTF8ToString(pointer: number): string;
  /** Allocates a decoder with one started session. Returns 0 on failure. */
  _x9_create(): number;
  /** Frees a decoder allocated by _x9_create. A zero pointer is a no-op. */
  _x9_destroy(context: number): void;
  /** Decodes one framed chunk. Returns the bytes written, 0 on failure. */
  _x9_decompress(
    context: number,
    source: number,
    sourceLength: number,
    destination: number,
    destinationCapacity: number,
  ): number;
  /** The decoder's own error text. Diagnostics only; never surfaced. */
  _x9_last_error(context: number): number;
  /** Allocates linear memory for the caller. Returns 0 on failure. */
  _x9_malloc(byteLength: number): number;
  /** Releases memory from _x9_malloc. A zero pointer is a no-op. */
  _x9_free(pointer: number): void;
}

declare const createXpress9Module: (
  options?: Xpress9ModuleOptions,
) => Promise<Xpress9Module>;

export default createXpress9Module;
