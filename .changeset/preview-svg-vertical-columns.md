---
'@pptx-kit/preview': minor
---

feat(preview): vertical text (`vert`, `vert270`, `eaVert`, `mongolianVert`,
`wordArtVert`) and multi-column bodies (`numCol`/`spcCol`) now render in the
pure-SVG text mode (`textLayout: 'svg'`) used by server-side rasterization.
Previously they fell back to horizontal single-column layout; now the server
output matches the browser (`foreignObject`) path: rotated line stacking for
vertical text and PowerPoint-style sequential column fill for multi-column
bodies. Browser rendering is unchanged.
