---
'pptx-kit': minor
---

feat: extend `TextFormat` with the remaining commonly-authored
`CT_TextCharacterProperties` (ECMA-376 §17.18.83) attributes:

- `strike` — `true` / `false` / `'sngStrike'` / `'dblStrike'`
- `spc` — character spacing in 1/100 pt
- `kern` — kerning threshold in half-points
- `baseline` — superscript / subscript offset as a unit fraction
- `cap` — `'none'` / `'small'` / `'all'`
- `highlight` — per-run background color

All round-trip through `setShapeRunFormat` / `getShapeRunFormat` and
flow through `getShapeRunFormatEffective`'s inheritance cascade. The
playground renderer honours each of them in the rendered HTML.
