---
'pptx-kit': minor
---

feat(site/playground): table cell borders honor `<a:prstDash>`. The
reader already surfaced the dash token; the renderer now projects it
to an SVG `stroke-dasharray` (scaled by the border's authored width).
Applies to every side, the top-left → bottom-right diagonal, and the
bottom-left → top-right diagonal.
