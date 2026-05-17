---
'pptx-kit': minor
---

feat: `getShapeEffectsEffective(pres, shape)` walks the layout →
master placeholder cascade for `<a:effectLst>`. Effect lists override
rather than compose (matching PowerPoint's behaviour), so the first
layer that supplies any effects wins. Playground uses it so
placeholder shadows / glows / soft edges inherited from the master
finally render on slides that don't repeat the effect list.
