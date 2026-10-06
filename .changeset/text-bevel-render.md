---
'@office-kit/pptx-preview': minor
'@office-kit/pptx': minor
---

The preview now draws text bevels (`<a:bodyPr>` `<a:sp3d><a:bevelT>`, as PowerPoint's Soft Bevel and Sharp Bevel WordArt write them): beveled glyphs are shaded with a light and dark edge sized from the bevel's width and height and lit from the light rig's direction, in shape text, table cells and the editor while text is being edited. The 12 bevel presets get different edge profiles and shiny materials add a highlight; this is a 2-D approximation, so the camera, extrusion and contour are not drawn. Text without a bevel renders exactly as before. New `getTableCellText3D` reads a table cell's text 3-D, like `getShapeText3D` does for shapes.
