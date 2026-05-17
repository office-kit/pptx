---
'pptx-kit': minor
---

feat: `getSlideLayoutBackgroundPatternFill(pres, layout)` and
`getSlideMasterBackgroundPatternFill(pres, layout)` complete the
pattern-background cascade. Slides reporting `'pattern'` can now
resolve the actual preset / colors by walking slide → layout →
master, paralleling the gradient / solid / blip cascades.
