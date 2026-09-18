# The committed XPress9 decoder

`xpress9.wasm` is Microsoft's reference XPress9 decoder compiled to WebAssembly,
and `xpress9.mjs` is the Emscripten loader that instantiates it.
`@consultchimps/pbi` needs them to decompress the Analysis Services backup image
inside a `.pbix` model part.

Both are build products and both are committed, which ADR 0004 Decision 2 asks
for: a contributor cloning the repository, and a consumer installing the
package, must get a working decoder without installing a C toolchain. The price
is that a reviewer cannot read them, so
`.github/workflows/wasm-reproducibility.yml` rebuilds them from the vendored
source and fails unless the rebuild is byte for byte what is committed here.

`xpress9.d.mts` is hand-written, not generated. It is the typed contract between
the loader and `src/xpress9/runtime.ts`, and it changes when the build script's
export lists change.

## Files

| File           | Bytes  | SHA-256                                                            |
| -------------- | ------ | ------------------------------------------------------------------ |
| `xpress9.wasm` | 41,380 | `502dea8357f428f0827dde03f8a519baa81c607238eefd68bad99370a016776a` |
| `xpress9.mjs`  | 5,483  | `f860c5f09ed15cc7f04c52ab84b0acc166937987163bcad0b1e5666a29c93c60` |

## Toolchain

Emscripten **6.0.9** exactly, and nothing else. The build script refuses to run
against any other version, because the artifact's bytes are a function of the
toolchain and a silent upgrade would turn the reproducibility check into a
failure nobody expected. CI runs the build inside `emscripten/emsdk:6.0.9`
pinned by immutable digest, recorded in the workflow.

A toolchain upgrade is therefore a single reviewed change that moves the version
in the build script, the digest in the workflow, the artifacts here and this
table together.

## Rebuilding

```sh
node packages/pbi/scripts/build-wasm.mjs
```

Run it from anywhere; it resolves its own paths and writes only `xpress9.mjs`
and `xpress9.wasm` into this directory, then prints the size and SHA-256 of
both. Nothing in `pnpm build` invokes it, so an ordinary build never needs a C
toolchain.

`emcc` has to be on PATH, or `EMCC` has to name it. On a machine with Docker,
the exact toolchain CI uses is

```sh
docker run --rm -v "$PWD:/src" -w /src \
  emscripten/emsdk@sha256:3ba391c5b1554e06f9af0a69652ff20919dff14e771619339dba988bd62574b5 \
  node packages/pbi/scripts/build-wasm.mjs
```

On a machine with no Docker, a local emsdk at 6.0.9 produces the same bytes, and
the check that matters is whether the hashes printed by the build match the
table above. CI is the authority either way: a passing local build is not a
reproducibility check.

## The flags, and why each one

Every flag is pinned in `packages/pbi/scripts/build-wasm.mjs`, which is the only
place the build is described. In short:

| Flag                                                                             | Why                                                                                                                                                                                                                                                                                                              |
| -------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `-Oz`                                                                            | Measured against `-O2` and `-O3` on this source: 41,380 bytes against 56,000 and 55,954, so roughly 26 percent smaller. It is also the faster of the two here, which is the opposite of the usual expectation and is why it was measured rather than assumed. See the note below.                                |
| `-sMODULARIZE=1 -sEXPORT_ES6=1`                                                  | An ES module factory, so the loader imports cleanly into a bundler, a worker and Node.                                                                                                                                                                                                                           |
| `-sENVIRONMENT=web,worker`                                                       | No Node-specific code in the loader at all. A Node caller hands the bytes in through `wasmBinary`, which `src/xpress9/runtime.ts` always does, so nothing is ever fetched by the glue.                                                                                                                           |
| no `-pthread`                                                                    | Single-threaded. No `SharedArrayBuffer`, so no cross-origin isolation requirement on any page that hosts the decoder.                                                                                                                                                                                            |
| `-sALLOW_MEMORY_GROWTH=1`                                                        | A model larger than the sample corpus needs more than the initial heap, and a fixed heap would have to be sized for the largest model on every model. On the nine samples the heap never grows past the 16 MiB initial size; the earlier spike build, which did not pin the initial size, doubled to 33,554,432. |
| `-sINITIAL_MEMORY=16777216`                                                      | 16 MiB. The LZ77 window is 4 MiB at the format's maximum, and the rest is the caller's per-chunk source and destination pair. This starts above the common case without charging small models for the largest one.                                                                                               |
| `-sFILESYSTEM=0`                                                                 | The decoder touches no files.                                                                                                                                                                                                                                                                                    |
| `-sASSERTIONS=0`, `-DNDEBUG`                                                     | No debug text, no debug code paths.                                                                                                                                                                                                                                                                              |
| `-sDYNAMIC_EXECUTION=0`                                                          | No `eval` and no `new Function`, so the module loads under a strict content security policy.                                                                                                                                                                                                                     |
| `-sSTRICT=1`                                                                     | A deprecated or misspelled setting is a build failure, not a silent no-op.                                                                                                                                                                                                                                       |
| `-sEXPORTED_FUNCTIONS`, `-sEXPORTED_RUNTIME_METHODS`, `-sINCOMING_MODULE_JS_API` | The module's JS surface is exactly six functions, two runtime helpers and two module options, pinned rather than inherited.                                                                                                                                                                                      |
| `-g0`, `-fno-ident`                                                              | No DWARF, no names section, no source map, no toolchain identity string.                                                                                                                                                                                                                                         |
| `-ffile-prefix-map=...`                                                          | The decoder's own consistency checks format `__FILE__` into their status text, so source paths reach the binary. The mapping keeps them relative and normalizes the include separator, which is `/` on Linux and `\` on Windows and would otherwise make the same source produce two different binaries.         |
| `SOURCE_DATE_EPOCH=0`                                                            | No wall-clock timestamp anywhere.                                                                                                                                                                                                                                                                                |

### The size and speed measurement

Decoding the 446,464-byte backup image of a public Microsoft sample, one fresh
module per run, ten timed runs after five warm-up runs, Node 24 on Windows:

| Build | wasm bytes | Mean decode |
| ----- | ---------- | ----------- |
| `-Oz` | 41,380     | 2.39 ms     |
| `-O2` | 56,000     | 6.77 ms     |

A fresh module per run is deliberate: it is the shape a real export has, one
module instantiated and one stream decoded, so the figure includes the engine
compiling the module rather than only the steady state of a long-lived one. A
smaller module is cheaper to compile, which is most of the difference. Read the
numbers as "`-Oz` costs nothing here", not as a claim that `-Oz` generates
faster code.

The build script also inspects what it produced and fails if any of that stops
holding: a `SharedArrayBuffer` or pthreads reference, a Node builtin, a source
map reference, a missing default export, an absolute path, or a path carrying a
native Windows separator.

## What is not checked here

The artifacts decode real Power BI output, and `test/xpress9-runtime.test.ts`
proves it against a chunk of a public Microsoft sample whose expected output
came from the reference decoder. What that test cannot prove is that the
committed bytes are what the vendored C compiles to. Only the workflow can, and
only in the pinned container.
