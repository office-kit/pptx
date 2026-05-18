---
'pptx-kit': minor
---

feat: `setTableCellAnchor(cell, 'top' | 'center' | 'bottom' | null)` and
`setTableCellMargins(cell, {left?, right?, top?, bottom?} | null)` —
writers for two `<a:tcPr>` properties that already had readers
(`getTableCellAnchor`, `getTableCellMargins`). The anchor setter maps
`top`/`center`/`bottom` to the schema's `t`/`ctr`/`b` values and clears
the attribute on `null`. The margins setter writes per-side EMU on
`marL`/`marR`/`marT`/`marB`; sides set to `null`/`undefined` are
stripped (PowerPoint falls back to its defaults); passing the whole
arg as `null` clears every side. Both create `<a:tcPr>` if absent.
