---
"@office-kit/pptx": minor
"@office-kit/pptx-preview": minor
"@office-kit/pptx-dev": minor
---

Slide numbers, as live fields rather than typed text.

`setShapeTextField(shape, type, { text })` writes an `<a:fld>` — the slide's
number, a date, a footer — replacing the shape's text body the way PowerPoint
writes one, and carrying the replaced text's formatting onto the field.
`addSlidePlaceholder(slide, type)` restores a single slot the layout reserves
(`sldNum`, `dt`, `ftr`, …), where `addMissingSlidePlaceholders` restores them
all. `getPresentationFirstSlideNumber(pres)` reads `<p:presentation
firstSlideNum>`.

The preview now substitutes `slidenum` fields with the slide's own position
instead of whatever number the file last cached, counting from the deck's
`firstSlideNum`. The editor gains a deck-wide slide-number switch in the slide
panel and an "Insert field" command under Insert ▸ Text, both in English and
Japanese.
