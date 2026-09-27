# Third-party licences

`@consultchimps/pbi` is licensed under the Apache License, Version 2.0. See the
[repository `LICENSE`](https://github.com/consultchimps/consultchimps/blob/main/LICENSE)
for the full text.

This package ships a WebAssembly binary compiled from vendored third-party C,
and a TypeScript port of a third-party kernel. The notices below apply to that
code and travel with the published package, so an npm consumer receives them
with the binary they are about to run. The repository root
`THIRD-PARTY-LICENSES.md` carries the same two notices for the workspace.

---

## XPress9 decoder (vendored C, compiled into `xpress9.wasm`)

Microsoft's reference XPress9 implementation, used to decompress the Analysis
Services backup image inside a `.pbix` model part.

|                    |                                                                                          |
| ------------------ | ---------------------------------------------------------------------------------------- |
| Upstream           | <https://github.com/Hugoberry/xpress9-python>                                            |
| Commit             | `2503e827f95187d65c67ed491b184bf852256db7`                                               |
| Licence            | MIT                                                                                      |
| In this repository | `packages/pbi/vendor/xpress9/` (source), `packages/pbi/wasm/xpress9.wasm` (build output) |

Every vendored C source and header carries
`Copyright (c) Microsoft Corporation.` and `Licensed under the MIT License.` in
its own header. The packaging repository that redistributes them is copyright
Igor Cotruta. The exact file list, and the encoder files deliberately left out,
are recorded in `packages/pbi/vendor/xpress9/README.md`.

```
Copyright (c) Microsoft Corporation.
Licensed under the MIT License.

MIT License

Copyright (c) 2025 Igor Cotruta

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## xmhuffman kernel (ported to TypeScript, no C compiled)

The Huffman kernel that decodes VertiPaq string dictionary pages. Only
`src/xmhuffman_kernel.c` is used, and it is ported to TypeScript rather than
vendored, so no C from this project is compiled into anything we ship. The
ported TypeScript file attributes its origin in its own header.

|                    |                                                 |
| ------------------ | ----------------------------------------------- |
| Upstream           | <https://github.com/Hugoberry/xmhuffman-cython> |
| Commit             | `2f908c56f29ae526b3b35772186f095f0a5eaaec`      |
| Licence            | MIT                                             |
| In this repository | ported into `packages/pbi/src/`                 |

```
MIT License

Copyright (c) 2026 Igor Cotruta

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
