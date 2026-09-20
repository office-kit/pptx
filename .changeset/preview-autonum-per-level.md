---
'@office-kit/pptx-preview': patch
---

fix: auto-numbered lists restarted at 1 after a nested list

The preview kept a single numbering counter and reset it on every indent-level change, so a top-level list with nested items between its entries rendered as `1. / a. / b. / 1. / 1.`. PowerPoint keeps one counter per level: the outer list continues (`2.`, `3.`) and only deeper levels reset when a shallower paragraph starts. A non-numbered paragraph still restarts its own level, and a different numbering scheme at the same level starts over.
