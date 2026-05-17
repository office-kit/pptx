---
'pptx-kit': minor
---

feat: full stroke read-back surface — `getShapeStrokeCap`,
`getShapeStrokeJoin`, `getShapeStrokeCompound` plus the existing
`getShapeStrokeDash` / `getShapeStrokeArrow`. Renderers now have
enough information to reproduce dashed outlines, rounded vs square
caps, miter vs bevel joins, and per-end arrow heads.

The playground composes `stroke-dasharray` from the preset dash
patterns (cadence multiplied by stroke width as PowerPoint does),
emits SVG `<marker>` defs for triangle / stealth / diamond / oval
arrowheads on connectors and shapes, and maps cap / join through.
