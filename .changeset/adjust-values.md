---
'pptx-kit': minor
---

feat: `getShapeAdjustValues(shape)` returns the `<a:prstGeom><a:avLst>
<a:gd name=… fmla="val N"/></a:avLst>` map (preset adjust-handle
values). Only literal `val` formulas are surfaced; computed formulas
(`pin`, `+-`, etc.) reference the preset's built-in guides and
aren't useful without them.

Playground reads `adj` on `roundRect` to project the authored corner
radius — previously every rounded rectangle painted at a hard-coded
18%. Other presets (callouts, arrows, etc.) can adopt the same getter
as their renderers grow.
