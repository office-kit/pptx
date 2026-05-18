---
'pptx-kit': minor
---

feat: `findShapesWithAnimation(slide)` — returns every shape on the
slide whose `getShapeAnimation` is not `null`. Pair to
`slideHasAnimations`. Useful for "which shapes on this slide actually
animate?" audits before exporting to a video pipeline that doesn't
honor PowerPoint's timing tree.
