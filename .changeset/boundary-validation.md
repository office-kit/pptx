---
'pptx-kit': minor
---

Validate authoring inputs at the API boundary so out-of-range values throw a
clear `RangeError` instead of silently emitting a schema-invalid `.pptx` that
PowerPoint marks corrupt and "repairs".

A generative schema-validation sweep surfaced a whole class of defects where a
caller-supplied number/string was serialized straight into a constrained
ECMA-376 attribute. These now reject (or, for GUIDs, normalize) at the boundary:

- Run formatting: `setShapeRunFormat` font `size` (ST_TextFontSize, 1..4000 pt)
  and `spc` (ST_TextPoint). It also now accepts the 3-digit hex shorthand for
  run `color` / `highlight`, matching `setShapeFill` / `setShapeStroke`.
- Tables: `setTableStyleId` / `addSlideTable` `styleId` (ST_Guid — a lowercase
  GUID from `crypto.randomUUID()` is now accepted and normalized to uppercase;
  a non-GUID string throws); `setTableCellBorders` `widthEmu` (ST_LineWidth);
  `setTableColumnWidth` / `setTableRowHeight` / `addSlideTable` `w`/`h`
  (ST_PositiveCoordinate); `setTableCellMargins` (ST_Coordinate32).
- Charts: bar/column `overlapPct` (ST_Overlap), `gapWidthPct` (ST_GapAmount),
  and series `lineWidthEmu` (ST_LineWidth).
- Animations / transitions: `setShapeAnimation` `durationMs` and
  `setSlideTransition` `advanceAfterMs` (xsd:unsignedInt); the transition
  `effect` token is validated against the spec's effect set (an empty or unknown
  string previously produced non-well-formed or schema-invalid XML).
- Connectors / strokes: `addSlideLine` and `setShapeStroke` `widthEmu`
  (ST_LineWidth) and `addSlideLine` endpoint coordinates.
- Text boxes: `setShapeTextColumns` `count` (ST_TextColumnCount, 2..16) and
  `gapEmu` (ST_PositiveCoordinate32); `setShapeTextMargins` insets
  (ST_Coordinate32); `setShapeTextBodyRotationDeg` (guards the ST_Angle overflow).
- Fills: `setShapePatternFill` `preset` is validated against ST_PresetPatternVal
  (and the `PatternPreset` type is now the exact token union, not `string`).
- Shape / image / chart geometry: `addSlideShape` / `addSlideTextBox` /
  `addSlideImage` / `addSlideChart` and `setShapePosition` / `setShapeSize`
  `x`/`y` (ST_Coordinate) and `w`/`h` (ST_PositiveCoordinate).
