---
'@office-kit/pptx': minor
---

`setShapePatternFill` now accepts partial settings, so changing a pattern preset or one color preserves the other color's theme reference and imported transforms. Invalid presets no longer remove the existing fill before throwing.
