---
'pptx-kit': minor
---

feat(site/playground): shapes with slide-jump click actions
(`<a:hlinkClick action="ppaction://hlinksldjump"/>`) render as
in-page hash anchors. The renderer resolves the target via
`getShapeClickAction` and emits `<a href="#slide-N">`; each slide's
`<li>` carries `id="slide-N"` so clicks scroll to the target slide.
Plain URL click actions render the same way as shape-level
hyperlinks (with `target="_blank"`).
