---
"@office-kit/pptx": minor
---

`resetSlidePlaceholderTextFormatting` and `resetSlideLayout` now reach
placeholders inside a group. A group scales and turns what is inside it; it
does not decide what font the text is in, so a grouped placeholder inherits its
layout's formatting and appearance like any other. Both functions count those
placeholders in what they return.

`resetSlidePlaceholderGeometry` still leaves them where they are, and so does
the geometry half of `resetSlideLayout`. The layout states a rectangle on the
slide, while a grouped shape's geometry is written in its group's coordinate
space — restoring it would either tear the shape out of the arrangement it was
grouped into, or invent a rectangle the layout never described.
