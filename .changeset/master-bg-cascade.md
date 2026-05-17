---
'pptx-kit': minor
---

feat: `getSlideMasterBackground(pres, layout)` returns the master's
`<p:bg>` (both `<p:bgPr>` and `<p:bgRef>` forms). Playground extends
its background fallback chain to slide → layout → master so brand
backgrounds authored on the master alone finally render on inheriting
slides instead of falling through to the theme's `light1`.
