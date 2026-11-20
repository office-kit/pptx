---
'@office-kit/pptx': patch
'@office-kit/pptx-dsl': patch
---

`setShapeTextFormat` without a range (and so a DSL `Text`'s `size`, `color`, `bold`, …) now also formats the shape's equations and their fallback text. Equations keep their Cambria Math face; the fallback text takes the font too. The DSL README describes writing equations through `paragraphs`.
