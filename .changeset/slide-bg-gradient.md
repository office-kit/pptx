---
'pptx-kit': minor
---

feat: `getSlideBackgroundGradientFill(slide)` returns the gradient
stops + path for slides with a `<p:bgPr><a:gradFill>` background.
Playground paints gradient slide backgrounds via the same projector
that handles shape gradients (linear / radial / rect / shape).
