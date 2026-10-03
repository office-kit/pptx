---
"@office-kit/pptx": minor
---

Insert several slides at one position with `addSlideAt` and distribute formatted text ranges to multiple shapes with `setShapeParagraphs`. Batch text copies preserve paragraph properties, fields and hyperlinks, including when a destination is also the source. Slide insertion rejects layouts from another presentation before changing the deck.

Read all paragraph inline elements in one pass by omitting the paragraph index from `getShapeParagraphElements`.
