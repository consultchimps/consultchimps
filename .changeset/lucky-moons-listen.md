---
"@consultchimps/messages": minor
---

Explain a PowerPoint template inspection in plain language. `formatHumanResult`
now recognises the `pptx.inspect-template` operation: it says what the slide
contains, states that nothing was created or changed rather than pointing at
output files that do not exist, and labels the inspection metrics — malformed
placeholder locations, placeholders outside a supported text shape, and
placeholders split across text runs — instead of printing their internal names.
