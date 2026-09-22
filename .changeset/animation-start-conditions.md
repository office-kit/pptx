---
"@office-kit/pptx": minor
---

`setShapeAnimation` accepts `start` and `delayMs`. `'withPrevious'` runs an effect alongside the one before it and `'afterPrevious'` once that one has finished, both off the click that started their predecessor, so several shapes can animate from a single click. As a slide's first effect neither has a predecessor to follow, so both run as the slide appears. Existing calls are unaffected: the defaults still write the click-triggered, zero-delay tree they wrote before, byte for byte.
