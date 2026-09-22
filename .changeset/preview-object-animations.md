---
'@office-kit/pptx-dev': minor
---

Play a slide's object animations in presentation mode. Clicks, the space bar, the arrow keys, the on-screen buttons and the presenter view all drive one click order: a slide's build is played through before the deck moves on, stepping back walks the build first, and a slide arrived at backwards is shown played out. An effect that runs with or after the previous one starts as the slide appears when nothing precedes it, a paragraph build reveals a paragraph per click, reduced motion keeps the order without the motion, and a slide that advances itself waits for the effects it has started.

The presenter view runs the same player over its own copy of the slide, resumed at the same point in the same stop, so it shows what the audience can see — including an effect still fading in and one still waiting on its delay.

Only what the deck states is played. An effect this library can read but not reproduce — an unknown preset, a target it does not model, a start condition or delay the tree does not state — keeps its place in the click order but is never approximated: nothing is chained onto an effect of unknown length, no click stop is invented, and everything such an effect touches is left exactly as it was drawn. A shape id that is drawn more than once is treated the same way, since nothing says which of them a timing tree means. The presentation controls say how many of a slide's animations are being left out, in English and Japanese.
