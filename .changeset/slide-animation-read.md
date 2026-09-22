---
"@office-kit/pptx": minor
---

Add `getSlideAnimations` to read a slide's animation effects in click order, with each effect's target (whole shape or paragraph range), preset, start condition, duration, delay and paragraph-build settings. Effects this library cannot author — custom presets, composite effects driving several shapes, interactive sequences — are reported rather than dropped, marked read-only so they are never silently rewritten.

`slideHasAnimations` no longer reports a slide whose only `<p:timing>` holds a video or audio clip's play controls as animated.
