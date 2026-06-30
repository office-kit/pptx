---
'pptx-kit': minor
---

Fix a batch of "generates but is schema-invalid / wrong" authoring bugs, add an
AI-agent authoring skill, and smooth several LLM-facing rough edges.

Correctness fixes (output is now schema-valid in these cases):

- Notes slides emitted a `<p:notesSlide>` root instead of the spec's `<p:notes>`,
  failing schema validation.
- Combining a run's text `color` with `highlight` emitted them out of order.
- Combining stroke dash / arrowheads / join, or a paragraph's bullet with
  `setParagraphSpacing`, or a table cell's fill with `setTableCellBorders`,
  emitted child elements out of their schema-mandated order.
- Table cells containing leading/trailing spaces, tabs, or newlines emitted an
  illegal `xml:space` attribute (and a newline now correctly splits a cell into
  multiple lines).
- `setSlideTransition({ effect: 'none' })` emitted an invalid `<p:none/>`; per-effect
  attributes (`direction`/`orientation`/`thruBlack`) are now only emitted on
  effects that allow them, and `direction` is validated against the effect's own
  value domain (e.g. `blinds` takes `horz`/`vert`, `push` takes `l`/`r`/`u`/`d`) —
  a mismatched pair like `{ effect: 'blinds', direction: 'l' }` now throws instead
  of emitting schema-invalid XML.
- Charts emitted `<c:marker>`, `<c:smooth>`, `<c:invertIfNegative>`, and
  `<c:trendline>` on series kinds that don't permit them (e.g. a trendline on a
  `pie`/`doughnut`/`radar` series, which `CT_PieSer`/`CT_RadarSer` reject), and
  `valueAxis` `min`/`max` in the wrong order.
- `setShapeImageBrightness` / `setShapeImageContrast` emitted `<a:lumOff>` /
  `<a:lumMod>`, which aren't valid `<a:blip>` children; both now write a single
  schema-valid `<a:lum bright/contrast>`.
- `setShapeGradientFill` ignored its documented `path` / `focus` options, silently
  downgrading radial/shape gradients to linear.
- `importSlide` could emit a duplicate `rId` when the source slide's layout
  relationship wasn't `rId1`.
- `setShapeAnimation` wiped any pre-existing `<p:timing>` on the slide (losing a
  template's authored animations); it now merges, so multiple shapes can animate.
- `compactPackage` / `readPackagePart` / `setMediaPartBytes` matched part names
  case-sensitively, unlike the rest of the package layer — a referenced image
  whose rel-target case differed could be wrongly deleted or missed.

Behavior change:

- `setShapeImageContrast` now takes a `[-1, 1]` offset (`0` = no change) instead
  of the previous `[0, 2]` multiplier, matching the underlying `<a:lum contrast>`.
- Unstyled connectors (`addSlideLine` without an explicit stroke) previously
  emitted no line style and rendered invisibly; they now carry a default
  `<p:style>` (`lnRef`/`fillRef`/`effectRef`/`fontRef`) so the line is visible.

Ergonomics:

- Colors accept the CSS-style 3-digit hex shorthand (`#f0a` → `FF00AA`).
- New `setParagraphLineSpacing(shape, p, { kind, value })` (the writer counterpart
  to the existing getter).
- `setTableCellBorders` accepts a partial border per side, so `{ color, widthEmu }`
  type-checks without spelling out `dash` (the read type `getTableCellBorders`
  returns stays strict — all fields populated).
- `addSlideShape` `textAnchor` is narrowed to the valid vertical anchors
  (`'t' | 'ctr' | 'b'`).

Docs:

- New `skill/SKILL.md` — a guide for driving pptx-kit from an AI agent (canonical
  calls, design rules, footguns, and a QA protocol), with a verified worked
  example.
