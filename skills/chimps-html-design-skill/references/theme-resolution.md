# Resolving a client theme

Run this only when the client supplied a workbook, deck or document to match.
With nothing supplied, use `assets/palette.css` and skip to the report line at
the bottom. Version 1 resolves one light palette and no dark counterpart.

Do it by hand, with an unzip tool and a text editor. Tooling for this is planned
in `@consultchimps/theme` and `@consultchimps/xlsx` and does not exist yet, so
do not call it.

## 1. Read the theme part

An `.xlsx`, `.pptx` or `.docx` is a zip. The theme part is normally:

| Source     | Part                    |
| ---------- | ----------------------- |
| Excel      | `xl/theme/theme1.xml`   |
| PowerPoint | `ppt/theme/theme1.xml`  |
| Word       | `word/theme/theme1.xml` |

A deck with several slide masters carries `theme2.xml` and beyond. Use the theme
of the master the client's own slides sit on.

Inside, `a:clrScheme` holds `dk1`, `lt1`, `dk2`, `lt2`, `accent1` to `accent6`,
`hlink` and `folHlink`. Each holds either `<a:srgbClr val="4472C4"/>` or
`<a:sysClr val="windowText" lastClr="000000"/>`. Take `val` for the first and
`lastClr` for the second. `a:fontScheme` holds `a:majorFont` for headings and
`a:minorFont` for body text, each with an `a:latin` element whose `typeface`
attribute is the font name.

## 2. Reject a stock theme

The shipped Microsoft defaults are not a client identity, so a stock theme
counts as no theme supplied: use `assets/palette.css` unchanged. Signs of one:

- the `a:theme` element is named `Office Theme`
- accent1 is `4472C4` (Office 2013 to 2021), `4F81BD` (Office 2007 to 2010), or
  `156082` (the Aptos theme shipped from 2024)
- accent3 is `A5A5A5`, the grey slot no designer chooses

## 3. Map the roles

| Theme role             | Token                  |
| ---------------------- | ---------------------- |
| `lt1`                  | `--surface`            |
| `lt2`                  | `--surface-sunken`     |
| `dk1`                  | `--ink`                |
| `dk2`                  | `--ink-secondary`      |
| `accent1` to `accent6` | `--cat-1` to `--cat-6` |
| `hlink`                | link colour            |

`--cat-7`, `--cat-8`, `--ink-muted` and `--rule` keep their default values: a
theme has only six accents, and a chart needing a seventh series has a bigger
problem than its palette. `folHlink` is unused in version 1.

## 4. Validate every role before use

Contrast ratio is `(L1 + 0.05) / (L2 + 0.05)` on the lighter and darker of the
two colours, where `L = 0.2126R + 0.7152G + 0.0722B` over channels scaled to 0
to 1 and linearised as
`c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4`.

Check, in this order:

1. The value parses as six hexadecimal digits. A `sysClr` without `lastClr`
   fails here.
2. `--surface` is lighter than `--ink`. If the theme inverts them, the client
   uses a dark deck: version 1 does not, so substitute both.
3. Ink roles clear 4.5:1 against the resolved surface.
4. Accent roles clear 3:1 against the resolved surface. A default slot that
   replaces a failing accent follows the labelling rule in `SKILL.md` when it
   sits under 3:1 itself, as slots 3, 4 and 5 do.
5. An accent whose highest sRGB channel minus its lowest is under 25 reads as
   grey and cannot carry identity in a chart. It fails.
6. Two accents that clear less than 1.5:1 against each other are one colour to
   the reader. Keep the earlier slot, fail the later one.

A role that fails any check is replaced by the same token's value in
`assets/palette.css`. Never nudge a client colour into passing: substitute it.

## 5. Resolve the font

Put the `a:minorFont` latin typeface first in the body stack and keep the
default stack behind it, so a reader without the font still gets a sane face.
Never add a `@font-face`, a font file, or a Google Fonts link to reach it. An
empty `typeface` attribute, or a name starting with `+`, means no preference:
keep the default stack.

## 6. Print the substitution report

Give the report in the reply, and paste it as an HTML comment above the palette
block in the deliverable. One row per substituted role:

```text
Theme: Contoso Corporate (xl/theme/theme1.xml)
accent3 A5A5A5  ->  #1baf7a   grey, channel spread 0 under 25
accent5 F2F5F7  ->  #e87ba4   contrast 1.08:1 against surface, under 3:1
Roles used as supplied: lt1, lt2, dk1, dk2, accent1, accent2, accent4, accent6
```

With no theme, or with a stock one, the report is one line:
`No client theme supplied, ConsultChimps default palette used.`
