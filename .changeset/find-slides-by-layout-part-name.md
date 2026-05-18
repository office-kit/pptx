---
'pptx-kit': minor
---

feat: `findSlidesByLayoutPartName(pres, layoutPartName)` — finds every
slide whose resolved layout part name matches `layoutPartName` (e.g.
`'/ppt/slideLayouts/slideLayout3.xml'`). Pair to the existing
`findSlidesByLayoutName` / `findSlidesByLayoutType`. Keyed on the
actual package path, so it's stable across template-name collisions
and PowerPoint UI locales.
