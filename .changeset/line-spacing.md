---
'pptx-kit': minor
---

feat: `getParagraphLineSpacing(shape, p)` returns the paragraph's
`<a:lnSpc>` as `{ kind: 'pct' | 'pts', value }`. Percent values come
through as a unit fraction (1.5 = 150%); point values are pt.

The playground projects both forms to CSS `line-height` per paragraph,
and uses the existing `getParagraphSpacing` to project spcBef / spcAft
to `margin-top` / `margin-bottom`. Text blocks now keep the vertical
rhythm the deck authored instead of falling back to a fixed line
height for everything.
