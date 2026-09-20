---
'@office-kit/pptx': patch
---

fix: `getParagraphAlignment` and `getTableCellAlignment` returned an invalid `algn` value from a file as if it were a valid alignment

Both getters cast the attribute straight to `ParagraphAlignment`, so a hand-edited or third-party deck with `algn="bogus"` handed `'bogus'` to typed code. They now parse the value, and anything outside `ST_TextAlignType` reads as `null` (unset), like the other alignment readers already did.

Their return type is now the new `ParagraphAlignmentToken` (`'l' | 'ctr' | 'r' | 'just' | 'dist' | 'justLow' | 'thaiDist'`) instead of the wider `ParagraphAlignment`. Runtime values are unchanged — these two getters have always returned the spec token, so `setParagraphAlignment(shape, 0, 'center')` reads back as `'ctr'` — but the old type also listed `'center'`, which let `getParagraphAlignment(…) === 'center'` compile although it can never be true. That comparison is now a type error; compare with `'ctr'`, or read the plain-English name from `getParagraphPropertiesEffective` / `getTableCellParagraphs`.
