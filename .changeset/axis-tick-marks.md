---
'pptx-kit': minor
---

feat: chart value-axis tick marks. `ChartSpec.valueAxisMajorTickMark`
and `categoryAxisMajorTickMark` carry `<c:majorTickMark val="in|out|
cross|none"/>`. The playground value-axis renderer draws short stubs
on the appropriate side of the plot edge (default `out` matches
PowerPoint's stock look); `none` suppresses them entirely.
