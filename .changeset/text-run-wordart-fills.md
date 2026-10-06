---
"@office-kit/pptx": minor
---

Support gradient and pattern fills on formatted text runs through `textFill`,
including theme color transforms and round-tripping native WordArt fills.
The existing `color` property remains the solid text color shorthand; passing
both `color` and `textFill` is rejected.
