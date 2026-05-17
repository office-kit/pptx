---
'pptx-kit': minor
---

feat: chart reader now recognises scatter, bubble, radar, stock, and
(2D / 3D) surface charts and degrades them to the closest already-
modelled kind so renderers paint something useful instead of the
"unsupported chart kind" placeholder. Scatter / bubble series read
their `<c:yVal>` channel; their `<c:xVal>` / `<c:bubbleSize>` are
not yet surfaced.
