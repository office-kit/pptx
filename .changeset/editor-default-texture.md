---
'@office-kit/pptx-dev': patch
---

Editor: choosing Picture or texture fill for a shape or slide background that has no picture now fills it with the default Papyrus texture, tiled as in PowerPoint, in one undo step, instead of opening a file chooser. Use Insert... to choose a picture file. A shape filled with a picture or texture now shows the Format Picture pane title. The Texture gallery now matches Mac PowerPoint: five textures per row at PowerPoint's size and spacing, no texture marked as the current fill, and no More Textures... in the Format pane (use Insert...; the ribbon's Shape Fill ▸ Texture keeps it). In Format Shape and Format Background, Texture is now a row with a right-aligned swatch ▾ button, and the gallery opens aligned to its right edge.
