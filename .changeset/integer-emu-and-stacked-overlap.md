---
'pptx-kit': patch
---

Fix corrupt files from fractional EMU and sideways-spreading stacked bar charts

- **Whole-EMU coordinates.** `inches` / `cm` / `mm` / `pt` / `emu` now round to integer EMU, and every shape / table / text-box / connector / chart offset is rounded on serialization. Floating-point drift from unit math (e.g. `3090672.0000000005`) previously reached `<a:off>` / `<a:ext>`, which is invalid `ST_Coordinate` (xsd:long) — PowerPoint flagged the file as corrupt and "repaired" it by zeroing the offending offsets, collapsing shapes to the slide origin.
- **Stacked bar/column charts** now emit `<c:overlap val="100"/>` by default (and for `percentStacked`). Without it PowerPoint draws each series in its own sub-slot so the "stack" spreads sideways across the category. An explicit `overlapPct` still wins; clustered charts are unchanged.
