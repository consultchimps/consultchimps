---
"@consultchimps/pptx": minor
---

Report a PowerPoint template inspection as a structured operation result.
`inspectPresentationOutcomeBytes()` in `@consultchimps/pptx/bytes` returns the
placeholder report together with the same `OperationResult` every completed
ConsultChimps operation returns: the slide's counts as metrics, no artifacts,
and one plain-language warning for each condition that would make a population
refuse the template — malformed placeholder braces, placeholders outside a
supported text shape, placeholders split across text runs, and a slide with no
usable placeholders at all. Identical templates and options produce an identical
result. The existing `inspectPresentationBytes()` is unchanged and still returns
the placeholder report on its own.

`planPopulatePresentationBytes()` now honours the `signal` its options already
accepted, and `inspectPresentationBytes()` accepts one. Both read whole packages
before they can answer, so a caller that has moved on — a page replanning after
a keystroke, or inspecting a different slide — can stop that work rather than
only discard its answer.
