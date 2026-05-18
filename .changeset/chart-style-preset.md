---
'pptx-kit': minor
---

feat: `ChartSpec.chartStyle` — round-trip the chartSpace-level
`<c:style val="N"/>` PowerPoint chart-style preset (1..48). Encodes a
curated combo of theme accent colors, gradients, effects, and font
sizes from the PowerPoint "Chart Styles" gallery. Read and written for
round-trip parity; pptx-kit's renderers don't interpret the preset
yet, but the field survives save/reload.
