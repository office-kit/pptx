---
'@office-kit/pptx-preview': patch
---

Fixed bar-chart category labels in the preview. A label longer than 14 characters was cut short with "…", and labels wider than the fixed left margin ran off the chart (and off the slide). The plot now narrows so the longest label fits, as presentation apps do, and a label that would take more than half the chart's width wraps onto more lines, breaking at spaces or between Japanese / Chinese characters and never starting a line with closing punctuation such as `）` or `%`. A bar chart's value-axis title now sits under the horizontal value axis and its category-axis title beside the labels; both used to be placed as if the chart were a column chart.
