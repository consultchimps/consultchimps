---
"@consultchimps/xlsx": minor
---

The table readers now read through the streaming reader consolidation uses (ADR
0006): `readWorkbookTables`, `readWorkbookWorksheets`,
`readWorkbookExcelTables`, `readWorkbookNamedRanges`, their byte twins, and the
input of a single-table split. Each input is parsed once instead of twice.

Values are read as before, with the changes consolidation already made. An error
cell holds its text, `#DIV/0!`, instead of Excel's internal number for it. A
carriage return before a line feed is kept. A formula's cached text is unescaped
once. A worksheet holding a declared date with no text is now read rather than
refused. Consolidation again refuses a worksheet whose row number or cell
reference is present but unreadable, as every other reader does. A worksheet
whose part cannot be found is refused by every reader, as consolidation already
did: an Excel Table or named range on it is no longer skipped silently, and a
sheet the workbook's relationships do not name is no longer guessed by position.
