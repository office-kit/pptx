---
'@office-kit/pptx-dev': minor
---

Editor ruler: indent and tab markers now also appear for rotated, flipped and vertical text. Rotated text is measured along its own lines as if the shape were unrotated, and vertical text (`vert`, `eaVert`, `vert270`, …) is measured on the vertical ruler. While you drag an indent marker or a tab stop, the text reflows immediately; the change is still saved as one undo step on release, and Escape restores the original layout. With several paragraphs selected, the ruler shows the first paragraph's markers; dragging moves each paragraph relative to its own indents, and moving a tab stop changes only the paragraphs that have it. While editing, decimal tabs align on the run language's decimal separator, such as `,` for German.
