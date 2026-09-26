# Charts

ECharts is the only chart library: no second library, no hand-rolled SVG chart,
no chart image.

## Inlining ECharts

1. Take `echarts.min.js` from an install of the `echarts` package
   (`node_modules/echarts/dist/echarts.min.js`) or from a downloaded release.
2. Paste its contents between an empty `<script></script>` pair in the
   deliverable. Do not use `<script src>`, not even to a CDN.
3. Record the version in a comment above the tag, so a later edit can match it.

The minified bundle is roughly 1 MB of text. That is the price of an offline
file, and it is worth it. If size becomes a real problem, build a custom bundle
with only the chart types used, and say so in the same comment.

Initialise with the SVG renderer: `echarts.init(el, null, { renderer: "svg" })`.
SVG prints and zooms cleanly, and the canvas renderer's output does not.

## Choosing the mark

| Question the chart answers | Mark                              |
| -------------------------- | --------------------------------- |
| Ranking across categories  | horizontal bar, sorted by value   |
| Change over time           | line, one line per series         |
| Part of a whole            | stacked bar, or a table of shares |
| Two measures per item      | scatter                           |
| One number                 | the number, in text, not a chart  |

## Mark rules

- No 3D: no `echarts-gl`, no isometric bars, no perspective, no bevel or shadow
  on a mark
- No pie chart beyond two slices: three or more become a sorted bar chart, and a
  donut is a pie
- Gridlines on the value axis only: one hairline per tick in `--rule`, solid,
  behind the marks, no category-axis gridlines, and no line on the value axis
  itself
- No gradients, no glow, no textured fills, no animated entrance in a
  deliverable that will be printed or screenshotted
- Bar value axes start at zero. Line value axes need not, but say so in the
  caption when they do not.
- Sort categories by value unless the category has its own order, such as months
  or a maturity scale.
- Label the units in the axis name or the caption, never only in the tooltip: a
  printed page has no tooltip.

## Colour

Assign `--cat-1` to the first series, `--cat-2` to the second, and so on down
the list in `assets/palette.css`. Set them on the option's `color` array in that
order.

Never cycle. When the series outrun the eight slots, the chart has too many
series: group the tail into an "Other" series, split the chart, or show a table.

Single-series charts use `--accent` for every bar or point. Do not colour a
single series by category: that spends the whole palette on a distinction the
reader can already see on the axis.

Highlight one category by giving it `--accent` and every other category
`--ink-muted`: two colours, one message.

Colour is never the only encoding. When two series sit in the same hue family,
or when the chart will be printed in grayscale, add a direct label, a dashed
line style, or a gap.

## Text inside the chart

- Inherit the page font: `textStyle.fontFamily` set from the body font stack
- Four series or fewer: direct labels at the end of each line or bar, no legend
- Five or more: a legend at the top, in palette order
- Axis and legend text in `--ink-secondary`, data labels in `--ink`
- Tooltips are an extra, not the story. `tooltip: { trigger: "axis" }` is enough
  for most charts.

## RTL charts

ECharts does not mirror itself. In an RTL document, set `yAxis.position` to
`"right"` for category charts and `xAxis.inverse` to `true` where the category
axis runs horizontally. Leave a time axis running left to right, and leave
numerals in Latin digits, which is what a finance reader expects.
