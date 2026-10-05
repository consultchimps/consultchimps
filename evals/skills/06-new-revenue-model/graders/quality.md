---
type: llm
target: trace
---

Pass when: there is an Assumptions sheet with one input per row and named ranges
that the calculations use; the growth rate is left blank and flagged Needs Input
rather than invented; calculations are formulas, not pasted values; newer
functions carry the _xlfn. prefix; creator and lastModifiedBy do not name a
library; the agent ran the bundled check_workbook.py lint or says why not; the
final message has the three handover headings Directly audited, Needs a decision
and Gaps.
