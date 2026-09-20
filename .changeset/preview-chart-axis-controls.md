---
"@office-kit/pptx-preview": patch
---

fix: chart previews ignored axis visibility and category tick settings (#351)

Hiding an axis line now preserves its labels and tick marks, and category axes use the requested major and minor tick marks. Hiding the category axis or its line also removes the zero baseline in column, bar, line, and area charts. Scatter and bubble previews now hide the entire corresponding axis, including its labels and ticks, when the value-axis or category-axis visibility setting requests it.
