---
'pptx-kit': minor
---

feat(site/playground): layout-type badge tooltip carries the
user-visible layout name. Hovering the small `obj` / `title` / etc.
badge now reveals `layout: <Name> (type: <token>)` from
`getSlideLayoutName(layout)`. Helps identify which authored layout
each slide is bound to without leaving the playground.
