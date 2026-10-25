---
'@office-kit/pptx-preview': patch
---

Fixed the preview giving column, line, area and combo charts a taller plot area than presentation apps do, which hid data labels that overlap in the real chart. Line labels above a point also sat on top of the marker, so a bar label and a line label that collide in LibreOffice were drawn apart. The automatic layout of bar, column, line, area and combo charts now keeps 2% of the chart's width and height clear on each side, as LibreOffice does, and leaves a category-label row sized to the label font; a line label above or below its point stays 1 mm clear of the marker. Charts render with a slightly smaller plot area than before; charts with a manual inner plot-area layout are unchanged.
