---
'pptx-kit': minor
---

feat: `ChartDataLabels.textStyle` — the default-run text style for chart
data labels is now read and written. `<c:dLbls><c:txPr><a:defRPr/>`
is parsed into `ChartTextStyle` (sizePt / bold / italic / color) and
emitted in CT_DLbls schema order (after `<c:numFmt>`, before
`<c:dLblPos>`). Both the chart-level `dataLabels` and per-series
`series[i].dataLabels` honor the field.
