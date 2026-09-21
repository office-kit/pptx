---
'@office-kit/pptx': minor
---

feat(charts): `ChartSpec` is now a union over `kind`, so a chart cannot be given a field it has no element for

Every chart kind carries only the fields its OOXML element defines. A
`scatterStyle` on a column chart, an axis title on a pie, `bar3DShape`
without `view3D`, a scatter series with no `xValues` — each of these used to
either throw at save time or serialize into nothing. They are now type errors
at the call site.

`getShapeChartSpec` and `SlideChartData.spec` return the new `ReadChartSpec`,
which keeps every field optional: a deck authored elsewhere can combine
fields no kind draws, and dropping them on read would lose data. Narrow one
back to the write side with `isChartSpec` before passing it to
`addSlideChart` / `setChartSpec`.

Two behaviour fixes came out of the change:

- A 3-D line chart no longer reports `lineMarkers: false` on read.
  `<c:line3DChart>` has no `<c:marker>` element, so the value was meaningless.
- Pie and doughnut charts no longer get the deck's body-text colour written
  onto axis-label styles they have no axes to show.
