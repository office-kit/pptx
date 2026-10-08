---
'@office-kit/pptx': patch
'@office-kit/pptx-preview': patch
---

Colors with a large `satMod` (or `satOff`) now resolve to the color the reference desktop app paints. `resolveDrawingColor` capped saturation at 100%, as ECMA-376 §20.1.2.3.27 describes, but the reference desktop app does not cap it. Because of that cap, the theme gradients that use `satMod` 300–350% came out too light and too gray in the preview. The Themed Style table backgrounds and the gradient fills of the built-in 2007-era themes are the visible cases: an orange accent was off by up to 70 levels per channel. They now match the reference desktop app's own exports within 2 levels.

`resolveDrawingColor` now also applies the `hue`, `sat`, `lum`, `red`/`green`/`blue` (with their `Mod`/`Off` forms), `gamma` and `invGamma` transforms, which it ignored before.
