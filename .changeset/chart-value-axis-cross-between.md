---
'pptx-kit': minor
---

feat: `ChartSpec.valueAxisCrossBetween` — controls whether the value
axis crosses the category axis *between* tick marks (the default for
bar/column/area) or *at* each tick mark (the default for line/scatter).
Maps to `<c:valAx><c:crossBetween val="between|midCat"/>`. Read by
chart-reader, written by chart-builder in the correct CT_ValAx schema
order (after `<c:crossesAt>`).
