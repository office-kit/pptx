---
'pptx-kit-site': patch
---

site(playground): project `ChartSpec.categoryAxisLabelAlign` into the
horizontal category-axis tick labels. Authored `<c:catAx><c:lblAlgn
val="ctr|l|r"/>` now overrides the rotation-derived default
text-anchor (useful for multi-line labels that the author wants flush-
left under each column). Vertical bar-chart layout keeps its existing
right-edge alignment, which is the only sensible default there.
