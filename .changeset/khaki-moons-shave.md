---
'@office-kit/pptx-preview': minor
---

fix: narrow the `@office-kit/pptx` peer range to `^0.20.0`

The previous `>=0.13.0` range accepted core versions this renderer no longer
matches, and it kept every future core release in range so the renderer was never
republished alongside a breaking one. `@office-kit/pptx` 0.13–0.19 are no longer
accepted; upgrade the core package alongside this one.
