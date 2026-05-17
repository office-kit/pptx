---
'pptx-kit': minor
---

feat: `getSlideColorMapOverride(slide)` returns the slide's
`<p:clrMapOvr><a:overrideClrMapping/>` token-remap, or `null` when the
slide inherits the master's color map. Returned as a plain `Record`
with the eight stable tokens (`bg1`, `tx1`, `bg2`, `tx2`, `accent1`-
`accent6`, `hlink`, `folHlink`) keyed to their override targets.
Useful for renderers that need to know when a slide reinterprets the
theme's color story.
