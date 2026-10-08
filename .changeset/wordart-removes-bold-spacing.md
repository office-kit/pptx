---
'@office-kit/pptx': minor
'@office-kit/pptx-editor': patch
---

`TextFormat.bold` and `TextFormat.spc` accept `null`, which removes the `b` / `spc` attribute so the run inherits its weight and character spacing again. `false` and `0` still write `b="0"` / `spc="0"`, which override an inherited bold or spacing.

The editor's WordArt Quick Styles now remove bold and character spacing that a preset does not set, as the reference desktop app does, instead of writing `b="0"` and `spc="0"`. Text whose list style or placeholder makes it bold or spaced keeps that look after applying such a preset.
