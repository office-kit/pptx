---
'@office-kit/pptx': minor
---

feat: author and read the complex-script typeface (`<a:cs>`), and align the blank deck's font scheme with Office's

- `TextFormat.fontComplexScript` sets and reads `<a:cs typeface="…"/>`, the third typeface slot next to `font` (`<a:latin>`) and `fontEastAsian` (`<a:ea>`). It works everywhere the other two already did: `setShapeRunFormat` / `setShapeTextFormat`, `setShapeParagraphs` and `setTableCellParagraphs` (runs and `endFormat`), and it comes back from `getShapeRunFormat`, `getShapeRunFormatEffective` (including `+mj-cs` / `+mn-cs` and the theme's complex-script fallback), `getShapeParagraphElements`, `getParagraphEndFormat` and `getTableCellParagraphs`.
- `ChartTextStyle.fontComplexScript` writes `<a:cs>` on a chart title / axis / data-label / legend style, and the chart reader returns it. `font` keeps filling the Latin and East Asian slots only, so a chart can name a complex-script face on the legend alone.
- The three slots stay independent: setting one never changes the other two. Rewriting a slot writes its `typeface` attribute alone, so a `pitchFamily` / `charset` the source file carried on that element is dropped — as was already the case for `<a:latin>` and `<a:ea>`.
- `createPresentation`'s blank deck now carries the standard per-script font list (`<a:font script="Thai" typeface="Cordia New"/>` and its 46 siblings) in both `majorFont` and `minorFont`, the same list Office's own default theme ships. It is what resolves a face for text in a script no run names one for. Every deck `createPresentation()` produces therefore has a larger `theme1.xml` than before — its bytes change — and text in a script with no explicit typeface now resolves through that list.
- Chart titles and axis titles no longer carry a `lang="en-US"` language tag on their run. The chart-level tag is still `ChartSpec.language` (`<c:lang>`).
