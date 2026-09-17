# Vendored XPress9 decoder

Microsoft's reference XPress9 implementation, reduced to the translation units
the decoder needs. `packages/pbi/scripts/build-wasm.mjs` compiles exactly these
files, plus our own `shim.c`, into `packages/pbi/wasm/xpress9.wasm`.

## Upstream

|            |                                                                                                                                                 |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| Repository | <https://github.com/Hugoberry/xpress9-python>                                                                                                   |
| Commit     | `2503e827f95187d65c67ed491b184bf852256db7` ("Bump to v0.3.8", 2025-11-30)                                                                       |
| Licence    | MIT, `LICENSE` beside this file, copied verbatim                                                                                                |
| Copyright  | Copyright (c) Microsoft Corporation, on every C source and header below; Copyright (c) 2025 Igor Cotruta on the packaging repository as a whole |

That repository vendors Microsoft's reference implementation, and every file
listed below carries the Microsoft copyright line and the MIT grant in its own
header. The repository root `THIRD-PARTY-LICENSES.md` records the same thing for
the repository as a whole.

## Files, with their upstream paths

Copied verbatim, with no edits to their content. Upstream writes them with CRLF
line endings and the repository's `.gitattributes` normalizes tracked text to
LF, which is the one difference from upstream's bytes. It changes nothing that
is built: compiling the LF copies and the CRLF originals produces the same
`xpress9.wasm` and the same `xpress9.mjs`, verified by building both.

| Here                            | Upstream path                   |
| ------------------------------- | ------------------------------- |
| `src/Xpress9DecHuffman.c`       | `src/Xpress9DecHuffman.c`       |
| `src/Xpress9DecLz77.c`          | `src/Xpress9DecLz77.c`          |
| `src/Xpress9Misc.c`             | `src/Xpress9Misc.c`             |
| `include/xpress.h`              | `include/xpress.h`              |
| `include/xpress9.h`             | `include/xpress9.h`             |
| `include/Xpress9Internal.h`     | `include/Xpress9Internal.h`     |
| `include/Xpress9Lookup.i`       | `include/Xpress9Lookup.i`       |
| `include/Xpress9Lookup.i2`      | `include/Xpress9Lookup.i2`      |
| `include/Xpress9Lookup.i3`      | `include/Xpress9Lookup.i3`      |
| `include/Xpress9Lz77Dec.i`      | `include/Xpress9Lz77Dec.i`      |
| `include/Xpress9ZobristTable.h` | `include/Xpress9ZobristTable.h` |
| `LICENSE`                       | `LICENSE`                       |

`shim.c` is not upstream. It is ours, Apache-2.0 with the rest of the
repository, and its header records that it mirrors the upstream
`src/Xpress9Wrapper.c` minus the encoder.

## Deliberately not vendored

The encoder never enters the wasm module, so its translation units and its
include fragments are absent and cannot be compiled in by accident:

- `src/Xpress9EncHuffman.c`
- `src/Xpress9EncLz77.c`
- `include/Xpress9Lz77EncInsert.i`
- `include/Xpress9Lz77EncPass1.i`
- `include/Xpress9Lz77EncPass2.i`

`src/Xpress9Wrapper.c` and `include/Xpress9Wrapper.h` are absent too: `shim.c`
replaces them with a decoder-only surface.

The Python packaging around the C (`setup.py`, `pyproject.toml`, `xpress9.pyx`,
`xpress9.pxd`, `xpress9.c`, `e2e_test.py`, `README.md`) is not vendored either.

## Updating

Changing anything here changes `packages/pbi/wasm/xpress9.wasm`. Bump the commit
hash in this file, rebuild with `packages/pbi/scripts/build-wasm.mjs`, refresh
the hashes in `packages/pbi/wasm/README.md`, and commit the new artifacts in the
same pull request. `.github/workflows/wasm-reproducibility.yml` fails if the
committed artifacts do not match a fresh build.
