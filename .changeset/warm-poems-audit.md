---
'@office-kit/pptx-preview': minor
---

feat: add `auditTextLayout` — detect text overflowing its box (はみ出し) and unintended soft wraps (段落ち)

`auditTextLayout(pres, options)` measures every shape's text with the same
layout engine the preview renders with and reports `overflow-x` / `overflow-y`
issues (plus opt-in `soft-wrap` reports via `reportSoftWraps`). Results carry
an `approximate` flag when widths were estimated rather than measured.

`buildFontkitMeasurer` (the `/node` entry) now accepts `{ fonts }` to register
the deck's actual font files by their authored family names; registered fonts
also serve as glyph fallbacks, and glyphs no font covers are estimated
per-character (CJK ≈ 1em) instead of measured against missing-glyph advances.
