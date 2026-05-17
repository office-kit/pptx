---
'pptx-kit': minor
---

feat: `getShapeFillEffective(pres, shape)` walks the layout → master
placeholder cascade when the shape's own fill is `'inherit'`. Returns
the first non-inherit fill found. Playground reaches for it as its
primary fill source so placeholder fills authored on the master /
layout finally show through.
