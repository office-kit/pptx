---
'@office-kit/pptx-dev': patch
---

feat: slide text in the preview can be selected and copied

The viewer rendered each slide inside a sandboxed `<iframe>` with pointer events disabled, so nothing on a slide could be selected. The slide now lives in a shadow root on the stage: the deck's SVG stays isolated from the viewer's DOM and CSS, but its text (XHTML inside `<foreignObject>`) is ordinary selectable content, and arrow-key navigation keeps working after clicking into a slide. In presentation mode a click still advances the deck unless it made a selection.
