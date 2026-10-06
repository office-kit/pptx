---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': patch
---

Preset shapes are now drawn from ECMA-376's own preset definitions instead of hand-written approximations. The preview, the editor canvas and its Shapes gallery icons showed many presets wrongly — hearts, lightning bolts, suns, moons, clouds, brackets and braces, bent and curved arrows, the equation shapes, stars with 7 to 32 points — and now match PowerPoint, including the lit and shaded faces of cubes, cans and curved arrows, adjust handles on every preset, and elliptical arcs in custom geometry (whose `arcTo` angles were read as parametric instead of visual angles). Text in a preset now wraps inside the rectangle the definition gives it, such as an ellipse's inscribed rectangle, as PowerPoint does.

- `@office-kit/pptx`: new `getPresetGeometry(preset, size, adjustValues?)` evaluates any preset's paths and text rectangle in the same form as `getShapeCustomGeometry`.
- `@office-kit/pptx-preview` (breaking): `shapeCustomTextRect(custom, extent)` is replaced by `shapeTextRect(shape)`, which reads custom and preset geometry alike, and `resolveTextBodyRect` now takes `(bounds, margins, region)` with the region from `shapeTextRect`.
