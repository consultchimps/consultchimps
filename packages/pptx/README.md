# @consultchimps/pptx

Local PowerPoint template inspection and population for ConsultChimps.

```ts
import { populatePowerPointTemplate } from "@consultchimps/pptx";

const result = await populatePowerPointTemplate({
  templatePath: "profile-template.pptx",
  workbookPath: "companies.xlsx",
  worksheet: "Companies",
  headerRow: 1,
  templateSlide: 1,
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

## Implementation boundary

The package edits the Open XML parts inside a `.pptx` archive with the
MIT-licensed `jszip` package already used elsewhere in ConsultChimps. It clones
the selected slide XML and its relationships, then changes only supported `a:t`
text nodes. This preserves the selected slide's existing shape and run
properties without coupling the public API to a presentation-generation library.

## Initial limitations

- Only placeholders entirely inside one text run in an ordinary PowerPoint text
  shape are replaced.
- A placeholder split across runs is detected and rejected before output is
  written. Re-entering the complete placeholder with consistent formatting in
  PowerPoint usually places it in one run.
- Placeholders in tables, charts, SmartArt, and other non-shape text locations
  are rejected.
- The output contains only generated copies of the selected template slide.
- Text is not resized or truncated. Review long values in PowerPoint.
- Images, charts, tables, conditional visibility, formulas, animations, and
  SmartArt are not populated or edited.
