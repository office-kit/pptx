---
'pptx-kit-preview': patch
---

fix: percentage pattern fills (`pct5`…`pct90`) now render at the requested
coverage. They were drawn as a sparse 1–4 dot grid that read far too light —
`pct50` looked like ~5% ink instead of a 50% screen. They now use an ordered
(Bayer) dither so the tone matches PowerPoint.
