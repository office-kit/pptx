---
'pptx-kit': minor
---

feat: `setShapeTextColumns(shape, { count, gapEmu? } | null)` — multi-
column writer pairing the existing `getShapeTextColumns` reader. Writes
`<a:bodyPr numCol="N" [spcCol="EMU"]/>`. Passing `null` clears both
attributes so the text body falls back to PowerPoint's default single
column. `count` must be `>= 2` (single column is the default — pass
`null` instead); the function throws otherwise.
