---
'pptx-kit': minor
---

Close a final batch of correctness defects a generative schema sweep surfaced,
where the writer emitted a `.pptx` PowerPoint marks corrupt:

- **XML-illegal control characters** in any text field (shape text, table cells,
  notes, chart titles/categories/series, hyperlink tooltip/URL, section names,
  comments) used to serialize raw, producing a non-well-formed part that
  corrupts the whole package. They are now rejected at serialization with a
  clear error; the XML-legal whitespace controls (tab / LF / CR) still pass
  through. (XML 1.0 forbids the other C0 controls outright — they cannot even be
  escaped as numeric references.)
- **Chart percentages** are now range-checked at the boundary: `gapWidthPct`
  (ST_GapAmount, 0..500 — the previous limit of 65535 let 501..65535 through),
  doughnut `holeSizePct` (ST_HoleSize, 1..90), and pie/doughnut
  `firstSliceAngleDeg` (ST_FirstSliceAng, 0..360).
- **Shape effects**: `setShapeShadow` `blurEmu`/`offsetEmu` and `setShapeGlow`
  `radiusEmu` are validated as ST_PositiveCoordinate (fractional rounds;
  negative / non-finite / over-max throws) instead of emitting an invalid value.
- **Scheme-color round-trip**: the read-back getters return `scheme:<token>`, but
  the setters rejected that string. `setShapeFill` / `setShapeStroke` /
  `setSlideBackground` now accept the `scheme:` prefix, so `setX(getX(...))`
  round-trips. (An unknown `scheme:` token still throws.)
- **`importSlide` / `mergePresentations`**: importing a slide that contains a
  chart left a dangling `r:id` (the chart frame referenced a relationship the
  imported slide no longer carried), producing a corrupt package. The orphaned
  graphic frame is now dropped, matching the documented "charts are not imported
  in v1" behavior.
- **`addSlideShape` presets**: the math-operator tokens were misspelled
  (`minus`/`mult`/`div`/`equal`/`notEqual`) and not in `ST_ShapeType`, so the
  shape silently vanished on open. They are now the spec names `mathMinus`,
  `mathMultiply`, `mathDivide`, `mathEqual`, `mathNotEqual` (plus `mathPlus`).
