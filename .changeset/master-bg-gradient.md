---
'pptx-kit': minor
---

feat: `getSlideMasterBackgroundGradientFill(pres, layout)` returns
the master's gradient background when `<p:bg><p:bgPr><a:gradFill>`
is authored. Completes the three-level bg cascade for gradient
backgrounds — slides can fall through slide → layout → master.
