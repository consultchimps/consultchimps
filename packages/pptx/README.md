# @consultchimps/pptx

Local PowerPoint template inspection and population for ConsultChimps.

```ts
import { populatePowerPointTemplate } from "@consultchimps/pptx";

const result = await populatePowerPointTemplate({
  templatePath: "profile-template.pptx",
  workbookPath: "companies.xlsx",
  headerRow: 1,
  outputPath: "company-profiles.pptx",
});
```

The operation reads `{{field_name}}` placeholders from ordinary text shapes on
one selected template slide. It creates one populated slide for every nonempty
Excel row below the header and returns a ConsultChimps `OperationResult`.
`overwrite: true` is required to replace an existing output. Neither input file
is changed.

Use `inspectPowerPointTemplate()` to discover valid placeholders and diagnose
malformed, split-run, or unsupported placements before population.

## Populate without a filesystem

`@consultchimps/pptx/bytes` runs the same population entirely in memory, which
suits browsers and other environments without a filesystem. The records come
either from an array you already hold or from workbook bytes:

```ts
import { populatePresentationBytes } from "@consultchimps/pptx/bytes";

const { result, outputs } = await populatePresentationBytes({
  template: { name: "profile-template.pptx", bytes: templateBytes },
  records: [{ client_name: "North", revenue: "10" }],
});
```

`inspectPresentationBytes()` and `planPopulatePresentationBytes()` mirror the
path-based inspection and plan. Output names are sanitized portable filenames,
never paths, and identical inputs produce byte-identical presentations.

`inspectPresentationOutcomeBytes()` reports the same inspection as the
structured `OperationResult` every completed operation returns — counts as
metrics, no artifacts, and one warning for each condition that would make a
population refuse the template:

```ts
import { inspectPresentationOutcomeBytes } from "@consultchimps/pptx/bytes";

const { inspection, result } = await inspectPresentationOutcomeBytes({
  name: "profile-template.pptx",
  bytes: templateBytes,
});
console.log(result.metrics.placeholderFields, inspection.placeholders);
for (const warning of result.warnings) {
  console.warn(warning);
}
```

## Implementation boundary

The package edits the Open XML parts inside a `.pptx` archive with the
MIT-licensed `jszip` package already used elsewhere in ConsultChimps. It clones
the selected slide XML and its relationships, then changes only supported `a:t`
text nodes. This preserves the selected slide's existing shape and run
properties without coupling the public API to a presentation-generation library.

## Initial limitations

- Placeholders in ordinary PowerPoint text shapes may span adjacent text runs;
  the replacement is written into the first run and the remaining placeholder
  fragments are cleared while preserving run formatting.
- The first template slide and first worksheet are selected by default. Use
  `templateSlide` or `worksheet` to select another one.
- Placeholders in tables, charts, SmartArt, and other non-shape text locations
  are rejected.
- The output contains only generated copies of the selected template slide.
- Text is not resized or truncated. Review long values in PowerPoint.
- Images, charts, tables, conditional visibility, formulas, animations, and
  SmartArt are not populated or edited.
