---
"@consultchimps/pbi": minor
---

Publish `@consultchimps/pbi`: read the model inside a `.pbix` file and export
every exportable table to a deterministic Excel workbook with a JSON manifest of
what was left out. The package ships the XPress9 WebAssembly decoder it needs
and its third-party notices, reserves every allocation against a configurable
byte budget before making it, and reports an unexpected fault inside a column
decoder as `PBI_COLUMN_DECODER_ERROR` rather than as damaged source data. Hosts
can now follow the export with `onProgress`.
