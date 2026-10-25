---
'@office-kit/pptx-preview': patch
---

Fixed the preview giving column, line, area and combo charts a taller plot area than presentation apps do, which hid data labels that overlap in the real chart. A chart with a bottom legend drew its plot about 30 px taller (at 1280 px wide) than LibreOffice, and line labels above a point sat on top of the marker, so a bar label and a line label that collide in LibreOffice were drawn apart. The automatic layout of bar, column, line, area and combo charts now follows LibreOffice's: 2% of the chart's width and height stays clear on each side, a bottom legend sits 1.85 mm above the chart's bottom edge with another 1.85 mm before the plot, the plot leaves half a line above its top value label and a category-label row sized to the label font, and a line label above or below its point stays 1 mm clear of the marker. Charts render with a slightly smaller plot area than before; charts with a manual inner plot-area layout are unchanged.
