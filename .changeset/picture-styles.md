---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': minor
---

Picture Styles, and Compress Pictures that labels pictures the way the reference desktop app does.

- New `setShapePictureStyle(picture, 'Metal Oval')` applies one of the reference desktop app's 28 built-in picture styles. It writes exactly the `p:spPr` markup the reference desktop app (Mac, 16.113) saves for that style (geometry, fill, border, effects and 3-D, with literal colors) and keeps the picture, its crop and its position. `getShapePictureStyle` returns the style a picture carries exactly, or `null`. New `BUILTIN_PICTURE_STYLES` lists the style names in the order of the reference desktop app's gallery.
- New `setShapeImageCompressionState` / `getShapeImageCompressionState` write and read the picture's `a:blip/@cstate` (`'print'`, `'screen'`, `'email'`, …) together with the `a14:useLocalDpi` extension the reference desktop app writes beside it.
- Preview: pictures now draw their effects (outer and inner shadow, glow, soft edge, reflection), their own fill and an approximation of their 3-D: the camera rotation as a flat projection and a top bevel as edge lighting.
- Fix (preview): a reflection faded the wrong way, strongest at its far edge, and ignored its end position; it now starts at the shape's edge and fades out by `endPos`. A soft edge blurred the whole shape; it now only feathers the outline.
- The editor's Picture Format ▸ Picture Styles gallery works: the 28 styles in the reference desktop app's order, drawn by the preview renderer, with the reference desktop app's English and Japanese names as tooltips and the applied style checked.
- The editor's Compress Pictures writes `cstate` for Print, On-screen and Email, as the reference desktop app does, and only replaces a picture's pixels when cropped areas are removed or the resampled picture is smaller. Picture Quality now opens the reference desktop app's menu: Compress Pictures... and Upscale Picture (unavailable: it uses the vendor's cloud AI service).
