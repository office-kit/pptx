---
"@office-kit/pptx": minor
---

`setShapeAnimation` accepts `start` and `delayMs`. `'withPrevious'` runs an effect alongside the one before it and `'afterPrevious'` once that one has finished, both off the click that started their predecessor, so several shapes can animate from a single click. As a slide's first effect neither has a predecessor to follow, so both run as the slide appears.

`'afterPrevious'` needs to know when the effect before it ends. On a slide whose timing states that in a form this library does not model — an effect that runs indefinitely or states no duration, one that repeats or is rescaled, or one that starts from another node rather than at a fixed offset — the call throws and leaves the slide's timing untouched rather than placing the effect at a guessed moment. `'click'` and `'withPrevious'` need no such measurement and still work there.

Existing calls are unaffected: the defaults still write the click-triggered, zero-delay tree they wrote before, byte for byte.
