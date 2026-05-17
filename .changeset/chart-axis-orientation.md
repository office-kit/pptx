---
'pptx-kit': minor
---

feat(chart): `ChartSpec.categoryAxisOrientation` and
`valueAxisOrientation` read `<c:catAx>/<c:valAx><c:scaling>
<c:orientation val="minMax|maxMin"/>`. Tools and renderers that
care about category render order (typically bar charts emit
`maxMin` so the first category sits at the top) can act on these
without dropping to XML.
