---
'pptx-kit': minor
---

feat: `setShapeHyperlink` and `setShapeRunHyperlink` now accept an
optional `tooltip` argument that writes a `tooltip="…"` attribute on the
emitted `<a:hlinkClick>`. Backwards compatible — calls that omit the new
arg behave exactly as before.

fix: `getShapeHyperlinkTooltip` previously only looked at the shape's
`<p:cNvPr><a:hlinkClick>`, missing the run-level tooltip that
`setShapeHyperlink` writes. It now scans run-level `<a:rPr>` first
(mirroring `getShapeHyperlink`'s read path) and falls back to the
shape-click hyperlink — so the reader / writer pair is consistent.
