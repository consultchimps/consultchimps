---
"@consultchimps/xlsx": minor
---

Add `describeWorkbookSource`, which describes a workbook read in pieces from a
random-access source such as `blobSource`, so a browser `File` is never held
whole. It gives the description `describeWorkbookBytes` gives for the same
bytes.
