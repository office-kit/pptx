---
'pptx-kit-site': patch
---

site(preview): render text as pure SVG `<text>` (not `<foreignObject>`) when a
`measureText` is supplied via the new `renderSlideSvg(pres, slide, opts)`
options, so the preview can be rasterized without a browser. The fidelity
harness uses this with a fontkit measurer + bundled metric-compatible fonts
(Carlito/Caladea/Liberation) and now scores text against LibreOffice ground
truth using a foreground-weighted SSIM. The browser playground is unchanged
(it keeps the `<foreignObject>` path by default).
