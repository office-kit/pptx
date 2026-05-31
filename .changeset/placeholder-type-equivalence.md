---
'pptx-kit': patch
---

fix: placeholder inheritance now applies the OOXML `ctrTitle`↔`title` and
`subTitle`→`body` type equivalence. A `ctrTitle` (centered title) now inherits
its layout/master `title` placeholder's `bodyPr` (e.g. `anchor="ctr"`),
`lstStyle`, and geometry instead of dropping them — fixing
`getShapeBodyPrEffective`, `getShapeBoundsResolved`,
`getShapeRunFormatEffective`, and `getParagraphPropertiesEffective` for
centered titles and subtitles.
