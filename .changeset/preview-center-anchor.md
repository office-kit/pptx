---
'pptx-kit-site': patch
---

site(preview): fix vertical alignment of center- and bottom-anchored text
(titles, most placeholders). The pure-SVG path positioned centered text ~4px
too high vs PowerPoint/LibreOffice; a calibrated `CENTER_ANCHOR_DROP`,
validated against ground truth, corrects it. Harness overall fg-SSIM ~0.46 →
~0.66 with no per-slide regression.
