---
'@office-kit/pptx-preview': patch
---

Paint an unstyled bullet in its paragraph's first-run colour. Without `<a:buClr>` anywhere in the cascade the renderer fell back to the deck's body-text colour, so a numbered agenda whose runs carry their own light colour drew its `1.` / `a.` markers in the default black and they sank into a dark background. PowerPoint and LibreOffice take the first run's colour, which is the same rule an un-sized bullet already followed for its size.
