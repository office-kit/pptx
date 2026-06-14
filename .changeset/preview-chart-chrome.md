---
'pptx-kit-preview': patch
---

fix: chart rendering now matches PowerPoint. The renderer drew an invented
light-gray chart-area frame, omitted axis spines, used faint inward tick stubs,
defaulted value-axis gridlines on, rendered every line/scatter marker as a
circle, and drew bar charts with the category axis upside-down. Now:

- The chart-area border is drawn only when the chart authors one.
- Value and category axes draw their spine and outward major tick marks.
- Major gridlines render only when authored (`<c:majorGridlines>`).
- Line / scatter / radar markers follow PowerPoint's automatic symbol
  rotation (diamond, square, triangle, x, …) when no symbol is authored.
- Bar charts order categories bottom-to-top, matching PowerPoint.
