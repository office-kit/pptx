---
'pptx-kit': minor
---

feat: chart builder writes back a wide slate of optional chart fields.
The chart-builder now emits, when authored on `ChartSpec`:

- `<c:varyColors>` (per chart kind), `<c:gapWidth>`, `<c:overlap>`
  on bar / column
- `<c:grouping>` honors `ChartSpec.grouping` (`'clustered' | 'stacked'
  | 'percentStacked' | 'standard'`) on bar / column / line / area
- `<c:dropLines>`, `<c:hiLowLines>` on line
- `<c:firstSliceAng>` on pie / doughnut; `<c:holeSize>` on doughnut
  honors `holeSizePct`
- `<c:majorGridlines>` (with optional `<c:spPr><a:ln><a:solidFill>`
  color), `<c:minorGridlines>` on the value axis
- `<c:title><c:overlay>` honoring `titleOverlay`

Round-trip test asserts the additions all survive read → save →
reload.
