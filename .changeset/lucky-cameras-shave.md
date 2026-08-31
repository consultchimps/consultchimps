---
"consultchimps": patch
---

Correct two over-broad claims in the packaged CLI. `sheets split --help` said it
copies "the complete workbook" once per value; it now says "the whole workbook"
and states that pivot tables and their caches are removed and reported as a
warning, which is what the operation has always done. The README said the CLI
requires Node.js 24 while `engines.node` declares `>=22.0.0`; it now says
Node.js 22 or later. The README's command overview also lists
`pptx inspect-template`, which it had never mentioned, and points at the CLI
reference rather than trying to be a full catalogue.
