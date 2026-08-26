---
"@consultchimps/xlsx": patch
---

Installing this package no longer requires network access to the SheetJS CDN.
SheetJS was declared as a runtime dependency pointing at a tarball URL, so every
`npm install` of `@consultchimps/xlsx` — and of the `consultchimps` CLI that
depends on it — had to reach `cdn.sheetjs.com`. Installs failed outright behind
registry-only allowlists, corporate proxies, private mirrors, and locked-down CI
runners.

SheetJS is now compiled into the published `dist` output instead, so the package
installs from the npm registry alone. Behaviour, the public API, and the
generated type declarations are unchanged; the published bundle is
correspondingly larger. The bundled Apache-2.0 code is attributed in the
package's `THIRD-PARTY-LICENSES.md`.
