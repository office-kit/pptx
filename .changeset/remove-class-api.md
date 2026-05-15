---
'pptx-kit': major
---

**BREAKING**: the class-based API (`Presentation`, `Slide`, `SlideShape`,
`SlideLayout`) has been removed. Use the free-function API for every
capability — one canonical path per operation.

| Was | Now |
| --- | --- |
| `Presentation.load(bytes)` | `loadPresentation(bytes)` |
| `Presentation.create()` | `createPresentation()` |
| `pres.save()` | `savePresentation(pres)` |
| `pres.slides` | `getSlides(pres)` |
| `pres.slideLayouts` | `getSlideLayouts(pres)` |
| `pres.addSlide({ layout })` | `addSlide(pres, { layout })` |
| `pres.removeSlide(slide)` | `removeSlide(pres, slide)` |
| `pres.moveSlide(slide, i)` | `moveSlide(pres, slide, i)` |
| `pres.duplicateSlide(slide)` | `duplicateSlide(pres, slide)` |
| `pres.replaceTokens(map)` | `replaceTokensInPresentation(pres, map)` |
| `slide.shapes` | `getSlideShapes(slide)` |
| `slide.findPlaceholder('title')` | `findSlidePlaceholder(slide, 'title')` |
| `slide.addTextBox(opts)` | `addSlideTextBox(slide, opts)` |
| `slide.addShape(opts)` | `addSlideShape(slide, opts)` |
| `slide.addImage(bytes, opts)` | `addSlideImage(slide, bytes, opts)` |
| `slide.addTable(opts)` | `addSlideTable(slide, opts)` |
| `slide.addLine(opts)` | `addSlideLine(slide, opts)` |
| `slide.setBackground(color)` | `setSlideBackground(slide, color)` |
| `slide.setTransition(opts)` | `setSlideTransition(slide, opts)` |
| `slide.setNotes(text)` | `setSlideNotes(slide, text)` |
| `slide.layout` | `getSlideLayout(slide)` |
| `slide.notes` | `getSlideNotes(slide)` |
| `slide.text` | `getSlideText(slide)` |
| `shape.text` | `getShapeText(shape)` |
| `shape.setText(value)` | `setShapeText(shape, value)` |
| `shape.position` | `getShapePosition(shape)` |
| `shape.setPosition(x, y)` | `setShapePosition(shape, x, y)` |
| `shape.setFill(color)` | `setShapeFill(shape, color)` |
| `shape.setStroke(opts)` | `setShapeStroke(shape, opts)` |
| `shape.setRotation(deg)` | `setShapeRotation(shape, deg)` |
| `shape.setHyperlink(url)` | `setShapeHyperlink(shape, url)` |
| `layout.name` | `getSlideLayoutName(layout)` |

Node entry (`pptx-kit/node`) drops the `Presentation` subclass; use
`loadPresentationFile` / `savePresentationToFile` instead.

**Why**: every capability used to have two paths through the public API
— a class method and a free function. The duplication hurt
discoverability (which one should you use?), made the bundle larger
(class consumers dragged the whole prototype in), and forced every
breaking change to land in two places. The free-function API is the
canonical surface from now on.
