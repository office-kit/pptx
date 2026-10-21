---
'@office-kit/pptx': patch
---

Fixed two problems with `ChartSeries.pointDataLabels`. A per-point label with literal `text` lost its `textStyle` (color, bold, size, font), because apps ignore a label's `<c:txPr>` once it carries its own text; the style is now also written into the label's rich text. And setting `pointDataLabels` without the series-level `dataLabels` threw at build time even though the types allow it; the other points of the series now keep the chart-level `dataLabels` (or show no label when there are none), so a per-point override no longer hides the labels of its neighbours.
