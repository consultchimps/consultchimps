#!/usr/bin/env node
/**
 * Builds packages/pbi/wasm/xpress9.mjs and packages/pbi/wasm/xpress9.wasm from
 * the vendored Microsoft XPress9 decoder in packages/pbi/vendor/xpress9.
 *
 * This is the one build command. It is what a maintainer runs locally and what
 * .github/workflows/wasm-reproducibility.yml runs twice in two clean copies, so
 * every flag is pinned here and nowhere else. Nothing in `pnpm build` invokes
 * it: the committed artifacts are what every contributor and every consumer
 * uses, which is the whole reason ADR 0004 Decision 2 commits them.
 *
 * Toolchain: Emscripten 6.0.9 exactly. The reproducibility workflow runs it
 * inside emscripten/emsdk:6.0.9 pinned by digest. Locally, put emcc on PATH
 * (emsdk_env) or set EMCC to its full path.
 *
 * Usage:
 *   node scripts/build-wasm.mjs [--out-dir <dir>]
 *
 * --out-dir defaults to packages/pbi/wasm. The script writes exactly two files
 * there and prints the SHA-256 of both.
 */

/* The repository's ESLint configuration declares no environment globals, so
   the one this script reads is named here rather than assumed. */
/* global process */

import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

/** The Emscripten version this build is pinned to, exactly. */
const REQUIRED_EMSCRIPTEN_VERSION = "6.0.9";

/**
 * Initial linear memory, 16 MiB.
 *
 * The decoder's own allocation is bounded: XPRESS9_WINDOW_SIZE_LOG2_MAX is 22
 * (include/xpress9.h:119), so the LZ77 window is 4 MiB, and the Huffman tables
 * and the decoder struct add well under another MiB. The rest is the caller's
 * per-chunk source and destination buffers. 16 MiB starts the module above the
 * decoder plus a typical chunk pair without growth, and ALLOW_MEMORY_GROWTH
 * covers the large fixtures, whose peak HEAPU8.byteLength the spike measured at
 * 33,554,432 bytes. A larger initial figure would charge every small model for
 * the largest one; a smaller one would force a grow on almost every model.
 */
const INITIAL_MEMORY = 16 * 1024 * 1024;

/**
 * -Oz, not -O2. Both were measured on this source, and the numbers are in
 * wasm/README.md: 41,380 bytes against 56,000, and a mean decode of 2.39 ms
 * against 6.77 ms on the same input. -Oz is smaller and, on the one-module
 * one-stream shape a real export has, faster too, because a smaller module is
 * cheaper for the engine to compile. Neither result was assumed, which is the
 * point of measuring: the usual expectation would have picked the other one.
 */
const OPTIMIZATION = "-Oz";

/** The three decoder translation units plus our shim. No encoder, ever. */
const SOURCES = [
  "vendor/xpress9/src/Xpress9DecHuffman.c",
  "vendor/xpress9/src/Xpress9DecLz77.c",
  "vendor/xpress9/src/Xpress9Misc.c",
  "vendor/xpress9/shim.c",
];

/** The complete JS-visible surface of the module. */
const EXPORTED_FUNCTIONS = [
  "_x9_create",
  "_x9_destroy",
  "_x9_decompress",
  "_x9_last_error",
  "_x9_malloc",
  "_x9_free",
];

/**
 * HEAPU8 so the loader can copy bytes in and out; UTF8ToString so a diagnostic
 * caller can read x9_last_error. src/xpress9/runtime.ts deliberately never
 * surfaces that text, but the entry point stays reachable.
 */
const EXPORTED_RUNTIME_METHODS = ["HEAPU8", "UTF8ToString"];

function parseArgs(argv) {
  let outDir = join(packageRoot, "wasm");
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === "--out-dir") {
      const value = argv[index + 1];
      if (value === undefined) {
        throw new Error("--out-dir needs a directory");
      }
      outDir = resolve(value);
      index += 1;
      continue;
    }
    throw new Error(`unknown argument: ${argv[index]}`);
  }
  return { outDir };
}

function run(command, args, options = {}) {
  // No shell. Several flags below carry a literal backslash, and a shell would
  // be free to eat it, which would silently drop the path normalization that
  // makes a Windows build and a Linux build produce the same bytes.
  const result = spawnSync(command, args, {
    cwd: packageRoot,
    encoding: "utf8",
    ...options,
    env: {
      ...process.env,
      // Reproducibility: no wall-clock timestamp may reach either artifact.
      SOURCE_DATE_EPOCH: "0",
      TZ: "UTC",
      LC_ALL: "C",
      ...options.env,
    },
  });
  if (result.error) throw result.error;
  return result;
}

function requirePinnedEmscripten(emcc) {
  const result = run(emcc, ["--version"]);
  if (result.status !== 0) {
    throw new Error(
      `could not run "${emcc}". Put Emscripten ${REQUIRED_EMSCRIPTEN_VERSION} on PATH, or set EMCC.\n${result.stderr ?? ""}`,
    );
  }
  const match = /\b(\d+\.\d+\.\d+)\b/.exec(result.stdout ?? "");
  const version = match?.[1];
  if (version !== REQUIRED_EMSCRIPTEN_VERSION) {
    throw new Error(
      `this build is pinned to Emscripten ${REQUIRED_EMSCRIPTEN_VERSION}; "${emcc}" reports ${version ?? "an unrecognized version"}. ` +
        "Changing the toolchain changes the committed artifacts, so it is a reviewed change to this script, the workflow digest and wasm/README.md together.",
    );
  }
  return version;
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

/**
 * Guards the properties the artifacts are required to have, so a flag that
 * silently stops applying is a build failure and not a runtime surprise.
 */
function assertArtifactShape(loaderPath, wasmPath) {
  const loader = readFileSync(loaderPath, "utf8");
  const failures = [];

  const forbiddenInLoader = [
    [
      "SharedArrayBuffer",
      "single-threaded build leaked a shared memory reference",
    ],
    ["pthread", "single-threaded build leaked a pthreads reference"],
    ["require(", "ENVIRONMENT=web,worker leaked Node-only module loading"],
    ["node:fs", "ENVIRONMENT=web,worker leaked a Node builtin"],
    ["sourceMappingURL", "a source map reference reached the loader"],
  ];
  for (const [needle, why] of forbiddenInLoader) {
    if (loader.includes(needle)) failures.push(`${why} (found ${needle})`);
  }

  if (!loader.includes("export default")) {
    failures.push("EXPORT_ES6 did not produce a default export");
  }

  // The decoder's own consistency checks format __FILE__ into their status
  // text, so source paths do reach the wasm. Relative ones are harmless and
  // useful. What would break the reproducibility check is an absolute path,
  // which names the directory the build ran in, or a native Windows separator,
  // which names the platform it ran on. Both are failures here rather than
  // surprises in CI.
  const wasmText = readFileSync(wasmPath).toString("latin1");
  const pathLike = /[\w.\-/\\]{4,}\.(?:c|h|i|i2|i3)(?![\w.])/g;
  const absolute =
    /^(?:[A-Za-z]:[\\/]|\/(?:home|root|tmp|Users|github|builds)\/)/;
  const rootSpellings = [packageRoot, packageRoot.split("\\").join("/")];
  for (const [name, text] of [
    ["loader", loader],
    ["wasm", wasmText],
  ]) {
    for (const spelling of rootSpellings) {
      if (text.includes(spelling)) {
        failures.push(`${name} embeds the build directory`);
      }
    }
    for (const [candidate] of text.matchAll(pathLike)) {
      if (absolute.test(candidate)) {
        failures.push(`${name} embeds the absolute path ${candidate}`);
      }
      if (candidate.includes("\\")) {
        failures.push(
          `${name} embeds the platform-shaped path ${candidate}; the separator normalization did not apply`,
        );
      }
    }
  }

  if (failures.length > 0) {
    throw new Error(`artifact check failed:\n  ${failures.join("\n  ")}`);
  }
}

function main() {
  const { outDir } = parseArgs(process.argv.slice(2));
  const emcc = process.env.EMCC ?? "emcc";
  const version = requirePinnedEmscripten(emcc);

  mkdirSync(outDir, { recursive: true });
  const loaderPath = join(outDir, "xpress9.mjs");
  const wasmPath = join(outDir, "xpress9.wasm");
  // Remove the previous outputs so a failed build cannot leave a stale artifact
  // behind for the comparison to pass against. The directory's other committed
  // files (README.md and xpress9.d.mts) are not build products and stay.
  for (const path of [loaderPath, wasmPath]) rmSync(path, { force: true });
  // Relative, POSIX-spelled, so nothing platform-shaped can reach the output.
  const outputArgument = relative(packageRoot, loaderPath)
    .split("\\")
    .join("/");

  const args = [
    OPTIMIZATION,
    // No entry point: the module is a library, there is no main().
    "--no-entry",
    // No DWARF, no names section, no source map, no toolchain identity string.
    "-g0",
    "-fno-ident",
    "-DNDEBUG",
    // The decoder's own consistency checks format __FILE__ into their status
    // text, so a handful of source paths do reach the binary no matter how it
    // is optimized. Three things keep them from making the build irreproducible.
    //
    // First, every source is named relatively on the command line below, so no
    // absolute path is ever handed to the compiler and the build directory
    // cannot appear. The map onto packageRoot catches anything that manages to
    // become absolute anyway.
    //
    // Second, the include fragments (the .i files) are opened through -I, and
    // the compiler joins the include directory to the file name with the
    // platform's own separator: "include/Xpress9Lz77Dec.i" on Linux and
    // "include\Xpress9Lz77Dec.i" on Windows. That single character is enough to
    // make a Windows build and a container build differ, so it is normalized
    // onto the forward-slash spelling. The mapping is passed on every platform,
    // so the flag list is identical everywhere and is a no-op where the
    // separator already matches.
    "-ffile-prefix-map=" + packageRoot.split("\\").join("/") + "=.",
    "-ffile-prefix-map=" + packageRoot + "=.",
    "-ffile-prefix-map=vendor/xpress9/include\\=vendor/xpress9/include/",
    "-I",
    "vendor/xpress9/include",
    ...SOURCES,
    "-o",
    outputArgument,
    // An ES module factory, so the loader is importable from a bundler, a
    // worker and Node alike.
    "-sMODULARIZE=1",
    "-sEXPORT_ES6=1",
    // Web and worker only. Node callers hand the wasm bytes in through
    // wasmBinary, so no Node-specific code is generated and a browser bundle
    // never sees a Node builtin.
    "-sENVIRONMENT=web,worker",
    // Single-threaded. No -pthread, so no SharedArrayBuffer and no cross-origin
    // isolation requirement on the page that hosts the decoder.
    "-sALLOW_MEMORY_GROWTH=1",
    `-sINITIAL_MEMORY=${INITIAL_MEMORY}`,
    // Emscripten's own default, pinned so the host's memory prediction in
    // src/xpress9/stream.ts reads a flag rather than a default a toolchain
    // upgrade could move. A heap that must grow grows by this fraction on top
    // of what was asked for, which is why summing payload bytes under-counts
    // the linear memory the module ends up holding.
    "-sMEMORY_GROWTH_GEOMETRIC_STEP=0.2",
    // Nothing in the decoder touches a file.
    "-sFILESYSTEM=0",
    "-sASSERTIONS=0",
    // No eval and no new Function, so the module loads under a strict CSP.
    "-sDYNAMIC_EXECUTION=0",
    // Refuse deprecated and misspelled settings rather than ignoring them.
    "-sSTRICT=1",
    `-sEXPORTED_FUNCTIONS=${EXPORTED_FUNCTIONS.join(",")}`,
    `-sEXPORTED_RUNTIME_METHODS=${EXPORTED_RUNTIME_METHODS.join(",")}`,
    // The only two module options the loader accepts, which keeps the rest of
    // emscripten's incoming API out of the glue.
    "-sINCOMING_MODULE_JS_API=locateFile,wasmBinary",
  ];

  const build = run(emcc, args);
  if (build.status !== 0) {
    process.stderr.write(build.stdout ?? "");
    process.stderr.write(build.stderr ?? "");
    throw new Error(`emcc exited with status ${build.status}`);
  }

  // Exactly two build products, and no third file emcc might have added (a
  // source map, a worker script, a symbol map) that would ship unnoticed.
  const allowed = new Set([
    "xpress9.mjs",
    "xpress9.wasm",
    "xpress9.d.mts",
    "README.md",
  ]);
  const unexpected = readdirSync(outDir)
    .filter((name) => !allowed.has(name))
    .sort();
  if (unexpected.length > 0) {
    throw new Error(
      `the build emitted unexpected files: ${unexpected.join(", ")}`,
    );
  }
  for (const path of [loaderPath, wasmPath]) {
    if (!existsSync(path)) throw new Error(`the build did not produce ${path}`);
  }

  // Emscripten writes the loader with the platform's line endings. Normalize
  // to LF so a Windows build and the Linux container build are byte-identical
  // and the committed file matches git's own LF checkout.
  const loaderText = readFileSync(loaderPath, "latin1");
  if (loaderText.includes("\r")) {
    writeFileSync(loaderPath, loaderText.replaceAll("\r\n", "\n"), "latin1");
  }

  assertArtifactShape(loaderPath, wasmPath);

  const lines = [
    `Emscripten ${version}, ${OPTIMIZATION}, INITIAL_MEMORY ${INITIAL_MEMORY}`,
    `xpress9.mjs   ${String(statSync(loaderPath).size).padStart(9)} bytes  sha256:${sha256(loaderPath)}`,
    `xpress9.wasm  ${String(statSync(wasmPath).size).padStart(9)} bytes  sha256:${sha256(wasmPath)}`,
  ];
  process.stdout.write(`${lines.join("\n")}\n`);
}

main();
