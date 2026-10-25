---
'@office-kit/pptx-dsl': minor
---

Added `fonts` to `Presentation`: `<Presentation fonts={{ minorLatin: 'Arial', majorLatin: 'Arial' }}>` sets the theme's body and heading typefaces (the same fields as `setPresentationFonts`). Text, tables and charts that name no font inherit them. Before, a deck could only set a font on each `Text`, so its charts kept the theme's Calibri while the slides used another font. Setting `textStyle.font` on each chart instead did not help Japanese or Chinese decks, because a chart's `textStyle.font` names one typeface for Latin and East Asian text alike, while the theme keeps its own East Asian and per-script fonts.
