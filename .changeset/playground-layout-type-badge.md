---
'pptx-kit': minor
---

feat(site/playground): show the slide's layout type (`title`, `obj`,
`twoObj`, `blank`, …) as a badge next to the slide title. Reads
`<p:sldLayout type="…">` via `getSlideLayout` + `getSlideLayoutType`
so deck audits can spot which layout each slide is bound to without
opening PowerPoint.
