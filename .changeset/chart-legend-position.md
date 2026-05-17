---
'pptx-kit': minor
---

feat(chart): `ChartSpec.legend` carries the `<c:legend><c:legendPos>`
token — `'r' | 't' | 'b' | 'l' | 'tr'`. Playground projects each
onto the appropriate edge (horizontal row for top / bottom, vertical
stack for the side / corner positions). Charts whose `<c:legend>`
sets `position` to `null` paint without a legend.
