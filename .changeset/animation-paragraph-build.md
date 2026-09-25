---
"@office-kit/pptx": minor
---

`setShapeAnimation` accepts `byParagraph`, which reveals a shape's text one paragraph at a time instead of animating the shape as a whole — PowerPoint's and Google Slides' "By paragraph". Each paragraph gets its own effect with the `start` you asked for, so the default advances a paragraph per click, and they share the one `<p:bldP build="p">` that makes PowerPoint treat them as a single build. `getSlideAnimations` reports each paragraph as its own step, targeting a paragraph range. A shape with no text is refused rather than silently animated as a whole, and a call that cannot be completed leaves the slide's existing timing exactly as it was.
