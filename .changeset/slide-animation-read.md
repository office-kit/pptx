---
"@office-kit/pptx": minor
---

Add `getSlideAnimations` to read a slide's animation effects in click order, with each effect's target (whole shape or paragraph range), preset, start condition, duration, delay and paragraph-build settings. Each effect also says which `<p:seq>` it belongs to, so a player can tell the slide's click order from a sequence a viewer triggers by clicking a shape, and carries separate `playable` and `editable` flags. Effects this library cannot author or render — custom presets, composite effects driving several shapes, interactive sequences — are reported rather than dropped, and flagged so they are never silently rewritten or folded into the normal progression.

`slideHasAnimations` no longer reports a slide whose only `<p:timing>` holds a video or audio clip's play controls as animated.
