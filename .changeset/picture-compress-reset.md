---
'@office-kit/pptx': minor
'@office-kit/pptx-dev': minor
---

Picture Format: Compress Pictures, Reset Picture & Size and the Artistic Effects gallery.

- New `getShapeImageArtisticEffect` reads the Artistic Effect PowerPoint applied to a picture or image fill (`a14:imgProps`, e.g. `'pencilSketch'`), or `null`. The embedded picture is already the effect's result, so the preview keeps drawing it as is; the effect, its JPEG XR original and the relationship to it survive edits, duplication and saving.
- Fix: setting a picture's transparency, brightness, contrast or recolor wrote the effect after the picture's `a:extLst`, which is schema-invalid; it now goes before it.
- The editor's Compress Pictures (Picture Format and File ▸ Compress Pictures...) offers PowerPoint's Picture Quality choices (High Fidelity, HD 330, Print 220, On-screen 150, Email 96 ppi, Use Original Quality), Delete cropped areas of pictures and Apply to. It downsamples PNG and JPEG pictures in the browser to the chosen resolution of their frame and removes cropped-away pixels, as one undo step.
- Reset Picture ▸ Reset Picture & Size now works: it also removes the crop and restores the picture's natural size at its own resolution.
- Artistic Effects shows PowerPoint's gallery (English and Japanese names) with the picture's current effect checked. Applying an effect stays unavailable: PowerPoint stores its own rendering plus a JPEG XR original, which the editor cannot produce. Picture styles stay unavailable too: PowerPoint's definitions are compiled into the application.
