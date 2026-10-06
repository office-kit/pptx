---
'@office-kit/pptx': minor
'@office-kit/pptx-dev': minor
---

Transitions can now carry a duration and a sound. `setSlideTransition` accepts `durationMs`, written as PowerPoint 2010 writes it (`p14:dur` inside `mc:AlternateContent`, with the nearest `speed` as the fallback), and `getSlideTransition` reads it back, including from decks saved by PowerPoint. New `setSlideTransitionSound` / `getSlideTransitionSound` embed a WAV sound or stop earlier sounds; changing the effect keeps the sound. The editor's Transitions tab gains Preview, Duration and Sound, and the preview plays fast/medium/slow transitions at PowerPoint's 0.5/0.75/1 s.
