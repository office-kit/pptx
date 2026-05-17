---
'pptx-kit': minor
---

feat: `getShapeTextColumns(shape)` returns `{ count, gapEmu? }` for
text bodies that author `<a:bodyPr numCol="N" spcCol="EMU"/>`.
Playground emits `column-count` / `column-gap` on the foreignObject,
so newspaper-style multi-column placeholders flow correctly.
