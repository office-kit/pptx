---
'pptx-kit-site': patch
---

chore(site): the preview renderer moved out of the site into a dedicated
`pptx-kit-preview` workspace package. The playground and REPL now import
`renderSlideToSvg` from `pptx-kit-preview`, and the fidelity harness renders
through `pptx-kit-preview/node` (`renderSlideToRgba`). No change to what the
site renders — this is a structural extraction so the Node "slide → image"
capability lives behind a clean, reusable module boundary (SVG in the browser,
PNG/RGBA in Node via resvg, no browser binary).
