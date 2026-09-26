---
name: chimps-html-design-skill
description:
  Standards for a client-ready HTML deliverable that opens offline, as one
  self-contained .html file with everything inlined, ECharts for charts, an
  ordered palette, a system font stack and RTL support. Use when producing or
  reviewing an HTML report, dashboard, one-pager or handout for a client, when
  charting in a web page, or when matching a client's Office theme.
license: MIT
metadata:
  palette: consultchimps-neutral-light
  version: "1"
  repository: consultchimps/consultchimps
---

# Client-ready HTML deliverables

## The delivery contract

One `.html` file. The client double-clicks it on a locked-down laptop with no
network and it works completely, first time.

- All CSS, JavaScript and data inline: no `<link>`, no `<script src>`, no
  `@font-face` URL, no `url(http...)` in CSS, no `<img src>` to a file or a host
- An image goes in as inline SVG or a `data:` URI.
- No `fetch`, no `XMLHttpRequest`, no web font, no analytics, no telemetry
- No build step: the file a text editor produces is the file the client opens
- Keep the data literal in its own `<script>` above the rendering code, so a
  figure can be corrected without reading the rest of the page.
- Include only the client data the deliverable actually shows.
- Print to PDF cleanly if you can: `break-inside: avoid` on figures and tables
  is most of it. It is a nice to have, not a requirement, and never a reason to
  cripple an interactive chart.

## Refuse on sight

3D charts, pie charts past two slices, Google Fonts or any hosted font, a
palette that cycles back to its first colour, and any external script or
stylesheet. Each has a section below saying why and what to use instead.

## Start from the shell

`assets/page-shell.html` is the skeleton: tokens, a `dir`-aware layout, the font
stack, and the two script placeholders. `assets/palette.css` is the palette,
pasted into the shell verbatim. Both are short enough to read before using.

## Colour

- One categorical palette per deliverable, the eight ordered slots in
  `assets/palette.css`
- Assign in order, first series to `--cat-1`, and never cycle back to the start.
  Series that outrun the palette are a signal to reduce series.
- Surface and ink roles carry every non-data colour: `--surface`,
  `--surface-sunken`, `--rule`, `--ink`, `--ink-secondary`, `--ink-muted`.
- Body text clears 4.5:1 against its surface, and a chart mark clears 3:1.
- Never reach past the tokens for a colour. A hex value written inline in a rule
  is a bug, because the theme resolution step cannot reach it.
- Version 1 ships one light palette, with no dark mode and no theme toggle.

## Fonts

A system font stack with a named preferred font at the front. The shell ships
`"Segoe UI", system-ui, "Helvetica Neue", Arial, sans-serif`.

- Never a Google Fonts link or any other hosted font: government and corporate
  networks block them, and the deliverable then renders in Times New Roman on
  the one screen that mattered.
- Version 1 embeds no font file: the weight is not worth it.
- Use `font-variant-numeric: tabular-nums` wherever figures line up in a column.

## Charts

ECharts, inlined, and no other chart library. The rules that decide whether a
chart is acceptable are in `references/charts.md`, and the short version is:

- No 3D, in any form
- No pie chart beyond two slices: three or more become a sorted bar chart
- Sober gridlines: value axis only, one hairline in `--rule`, no category-axis
  lines, no axis line on the value axis
- Palette slots in order, never cycled
- Colour is never the only encoding: two series in one hue family need a direct
  label or a line style too.

## RTL

Capable, not primary. It costs a handful of rules, so pay them up front:

- Set `dir` on `<html>`, never on individual elements.
- Use logical properties everywhere: `margin-inline-start`, `padding-inline`,
  `border-inline-start`, and `text-align: start` and `end`, never `left` or
  `right`.
- Swap in an Arabic-capable font stack under `[dir="rtl"]`, as the shell does.
- Mirror charts explicitly, because ECharts does not: see the RTL section of
  `references/charts.md`.
- Test by setting `dir="rtl"` on the shipped file and looking at it once.

## Matching a client theme

When the client supplies a workbook or deck to match, read its Office theme,
validate every colour role, substitute a neutral default for each role that
fails, and print a substitution report saying which roles were substituted and
why. A stock Office theme counts as no theme supplied: use the default palette.
The whole procedure, done by hand with an unzip tool and a text editor, is in
`references/theme-resolution.md`.

## The page itself

- A title, the client name, the date, and the scope in one sentence, at the top
- Every figure gets a caption that says what it means, not what it plots.
- A source and method note at the bottom, because a consultant will be asked
- No placeholder text, no lorem ipsum, no invented figure
- A number that is not in the data does not go on the page.

## Before handing it over

1. Search the file for `src="http`, `href="http`, `url(http`, `fetch(` and
   `googleapis`. Every hit is a broken deliverable.
2. Open it with the network disabled and confirm every chart draws.
3. Confirm the palette slots run in order and none repeats.
4. Set `dir="rtl"`, reload, confirm the layout mirrors, then set it back.
5. Print to PDF and skim the pages.
6. State the file size in the handover message if it is above 5 MB, because a
   mail gateway may not carry it.
