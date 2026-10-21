---
'@office-kit/pptx': minor
---

New `setChartData(chart, { kind?, categories?, series? })` replaces the data of a bar, column, line, area, pie or doughnut chart. When the kind and the set of series names stay the same, it rewrites only the cached values and their ranges in the chart part, so formatting from another app (colors, number formats, layout, extensions) is kept; changing the kind or the series builds the chart again through `setChartSpec`, carrying over its title, axis titles, legend and series colors. The embedded workbook is rewritten either way. It throws when `categories` or `series` is empty, when a series does not have one value per category, and on charts of other kinds; a date category axis takes its categories as date serial numbers.
