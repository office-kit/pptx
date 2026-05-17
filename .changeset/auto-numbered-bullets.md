---
'pptx-kit': minor
---

feat(site/playground): render auto-numbered bullets. Paragraphs with
`bulletStyle === 'number'` or `{ autoNum: '…' }` now emit the next
number in sequence (1., 2., 3., …; A., B., C., …; i., ii., iii., …)
rather than a generic dot. Counter resets on a non-numbered paragraph
or a level change, matching PowerPoint's behaviour.

Covers the common `ST_TextAutoNumberScheme` tokens — arabicPeriod /
ParenR / ParenBoth, romanUc / Lc with Period / ParenR / ParenBoth,
alphaUc / Lc with Period / ParenR / ParenBoth.
