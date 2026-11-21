---
'@office-kit/pptx': minor
---

Read shapes wrapped in `mc:AlternateContent` in a shape tree. The desktop app saves a shape holding an equation this way, and such shapes were left out of `getSlideShapes`, so the preview drew nothing for them. The reader now takes the branch an application without newer features would take: the a14 choice (the equation) or, for any other choice, the fallback. Removing, reordering, grouping, ungrouping and copying such a shape move the whole wrapper.

`getSlideShapes` can now return more shapes for decks that use this markup, which shifts the index of later shapes.
