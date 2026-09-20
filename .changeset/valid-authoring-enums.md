---
'@office-kit/pptx': minor
---

Validate enum inputs at authoring boundaries and throw a descriptive `RangeError`
instead of writing invalid PowerPoint XML or silently selecting another mode.

Affected APIs:

- Shapes: `addSlideShape` preset and text anchor, including calls through the TSX DSL.
- Strokes and fills: `setShapeStrokeDash`, `setShapeStrokeCap`, `setShapeStrokeJoin`,
  `setShapeStrokeCompound`, `setShapeStrokeArrow` (end, type, width, length), and
  `setShapeGradientFill` path.
- Text: `setShapeAlignment`, `setParagraphAlignment`, `setShapeTextAnchor`,
  `setShapeTextAutoFit`, `setShapeTextWrap`, `setShapeTextDirection`, `setShapeBullets`,
  `setParagraphBullet`, and `setShapeText` bullets; `setShapeTextFormat`,
  `setShapeRunFormat`, and `setShapeParagraphs` formatting and alignment.
- Tables: `setTableCellAlignment`, `setTableCellTextFormat`, `setTableCellParagraphs`,
  `setTableCellTextDirection`, `setTableCellAnchor`, and `setTableCellBorders` dash.
- Charts: `addSlideChart` and `setChartSpec` kind, grouping, axes, legend, layout,
  data labels, series markers, line dash, trendlines, error bars, and chart styles.
- Slides: `setSlideSize` type, `setSlideTransition` speed and split orientation,
  and `setShapeAnimation` effect.

Previously accepted inputs such as preset `'rounded-rect'`, alignment `'middle'`,
stroke dash `'dotted'`, text auto-fit `'shrink'`, and auto-numbering `'bogus'` now
throw `RangeError`. Underline and strike strings must be exact XSD tokens:
use `'sng'` / `'sngStrike'` (or `true`), not `'single'`. Existing JavaScript
`null` values for underline, strike, and cap still remove the corresponding
attributes; their TypeScript types are unchanged.

Unknown chart kinds now throw `RangeError: addSlideChart: kind: … is not one of: …`
instead of `Error: unsupported chart kind: …`; `setChartSpec` uses its own API name
in the new message. Code that checks the exception class or message should adapt.

Valid ECMA-376 shape presets remain supported, including presets outside the
TypeScript autocomplete list. Rejected bullets and transitions preserve existing
content, and rejected table-cell alignment or text-format enums do not create a
text body.
