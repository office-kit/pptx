---
'pptx-kit': minor
---

feat: `getShapePatternFill(pres, shape)` returns the pattern preset
token plus the foreground / background colors resolved against the
deck's theme. Pairs with the existing `setShapePatternFill`. The
playground renderer now paints real SVG `<pattern>` tiles for the
common `ST_PresetPatternVal` tokens (pct5..pct90, light/dark diagonal
and horizontal/vertical stripes, grids, weave, wave, sphere, diamonds)
instead of substituting a flat tint.
