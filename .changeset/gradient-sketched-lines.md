---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': minor
---

Gradient lines, sketched lines and more text paints:

- **Gradient lines.** `setShapeStroke(shape, { fill: { kind: 'gradient', ...gradient } })` writes `<a:ln><a:gradFill>` with the same gradient options as `setShapeGradientFill`, and a text outline takes the same `fill` (`setShapeTextFormat(shape, { outline: { fill } })`). `getShapeStrokeGradient(shape, pres?)` reads it back; `getShapeStroke` / `getShapeStrokeEffective` now report `{ kind: 'gradient' }` for such lines (previously `inherit`), so a `switch` over `ShapeStroke['kind']` needs a `gradient` case.
- **Sketched lines.** `setShapeStrokeSketch(shape, 'curved' | 'freehand' | 'scribble' | null)` / `getShapeStrokeSketch(shape)` write and read the reference desktop app's Sketched style (`ask:lineSketchStyleProps`). Other `<a:ln>` extensions are kept, and a sketch the reference desktop app saved (which replaces the geometry with the hand-drawn path) gets its original geometry back when the sketch is changed or removed.
- **Text fills.** `TextFormat.textFill` accepts `{ kind: 'none' }` (`<a:noFill/>`) and `{ kind: 'image', bytes }` (a PNG, JPEG, GIF, BMP, TIFF or WebP stretched over the text, as the reference desktop app writes it), on shapes, table cells and notes.
- **Keep text flat.** `setShapeTextFlat(shape, flat)` / `getShapeTextFlat(shape)` write `<a:bodyPr><a:flatTx/>`.
- **Preview** strokes gradient outlines on shapes, connectors and text, paints no-fill and picture-filled text, and draws a hand-drawn approximation of sketched outlines.
- **Editor**: Gradient line in Line and Text Outline (with the gradient type, direction, angle and stop controls of Gradient fill), the Sketched style menu (None, Curved, Freehand, Scribble), No fill and Picture or texture fill for text, and Keep text flat are now enabled. Text Effects ▸ Soft Edges stays disabled, as in the reference desktop app. The Text Fill gradient and pattern options no longer fail to apply.
