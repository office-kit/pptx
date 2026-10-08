---
'@office-kit/pptx': minor
---

The emphasis effects' own Effect Options:

- **Spin** takes `spinDirection` (`'clockwise'` / `'counterclockwise'`) and `spinDegrees` (90, 180, 360, 720 or any positive angle); **Grow/Shrink** takes `scaleDirection` (`'both'` / `'horizontal'` / `'vertical'`) and `scalePercent`; **Transparency** takes `transparencyPercent`; and Fill, Font, Line, Brush and Object Color, Color Pulse and Grow With Color take `color` — a theme slot or `#RRGGBB`, optionally with colour transforms (`AnimationColor`). `setShapeAnimation` and `updateSlideAnimation` write them, and `getSlideAnimations` reads them back on `SlideAnimationStep`.
- A spin through any angle, either way, is now read as `spin` with its `spinDegrees` and `spinDirection`. Before, only a full clockwise turn was named. A scale that stretches the two axes by different amounts, or a colour effect with two different colours, is still listed with `effect: null`.
- **Changed:** `updateSlideAnimation` with a new `effect` now gives it that preset's default duration, as picking a preset from the reference desktop app's gallery does. Before, it kept the duration the old effect had. Pass `durationMs` in the same patch to keep a length of your own. Changing only an option or the text build still keeps the duration.
