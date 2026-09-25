---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": minor
---

Edit a slide layout, not just apply one: `setSlideLayoutName`,
`setSlideLayoutBackground`, `clearSlideLayoutBackground` and
`setSlideLayoutPlaceholderBounds`.

A layout could be read and applied but never changed, so a deck's shared design
was fixed at whatever the template shipped. A layout handle now carries its own
document and writes back into the layout part, and every slide on the layout
follows the change — except where a slide set its own position or background,
which still wins.

Adding or removing a layout, adding a placeholder slot, and editing the slide
master are still not supported.

The editor exposes the settings through its properties panel and the
Design ▸ Layout ribbon group, in English and Japanese.
