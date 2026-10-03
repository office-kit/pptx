---
"@office-kit/pptx": minor
---

Animations on a slide can now be edited, not only added. `updateSlideAnimation` changes an effect's preset, duration, delay, start condition or paragraph build; `removeSlideAnimation` takes one out; `moveSlideAnimation` changes its place in the click order. Every effect is addressed by the `<p:cTn id>` `getSlideAnimations` reports, and that id survives each call — including a `byParagraph` toggle, where the addressed effect keeps its place and the rest of the build is new. Reordering and retiming recompute the offsets that make `afterPrevious` mean it, so the effects after the one you changed still start when the one before them ends, and an effect that starts as the slide appears is never quietly turned into one that waits for a click.

Copying a shape now brings its animations across as the effects they are: a build with a paragraph deleted stays deleted, per-paragraph timings survive, and a preset this library can only read comes along rather than being dropped. Removing a shape takes its effects with it and moves up whatever started when they ended.

An edit, copy or removal this library cannot reproduce faithfully is refused before anything changes, with a message naming what is in the way — a `<p:tn>` condition that would be left waiting on nothing, an effect shared with a shape that is staying behind, an interactive sequence, or a main sequence laid out in a way it could not write again. In particular, `copyShape` refuses before adding the shape, its relationships or its parts, and `removeShape` refuses before taking the shape out, so the presentation is left exactly as it was.
