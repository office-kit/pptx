---
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': patch
---

Paint text runs with a gradient fill (`<a:gradFill>`) or pattern fill (`<a:pattFill>`), such as PowerPoint's gradient and pattern WordArt presets. These runs were previously drawn in the default text color. A gradient spans the whole text block, including every line, as in PowerPoint. Theme colors and their tints are resolved. Pattern fills use the same tiles as shape pattern fills. This works in both the SVG and the browser (`foreignObject`) text layouts. The editor canvas now shows these fills, also while the text is being edited.
