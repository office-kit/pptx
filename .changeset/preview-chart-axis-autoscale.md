---
'@office-kit/pptx-preview': patch
---

Fixed chart value axes in the preview. The automatic range now follows the spreadsheet rule that presentation apps use: each open end is padded by 1/20 of the data range before rounding out to a major unit, and zero is left out of the range when the smallest value is at least 5/6 of the largest. A combo chart's secondary axis with rates of 50–53.4 was drawn 0–60, which flattened its line; it is now 49–54. The secondary axis also uses its own authored `<c:scaling>` (min, max, major unit, number format, display units), tick marks, gridlines and line color instead of ignoring them, its spine is drawn on the right edge of the plot rather than over the left one, and value-axis labels start past an outward tick mark so the tick no longer reads as a minus sign.
