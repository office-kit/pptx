---
'pptx-kit': patch
---

Numbered lists now carry a bullet font and explicit start number

`setShapeBullets('number')` / `setParagraphBullet(..., 'number')` emitted only
`<a:buAutoNum>`, with no `<a:buFont>` — so the auto-number glyph fell through to
whatever font happened to apply, instead of the theme's major font that
PowerPoint and PptxGenJS use. Numbered lists now emit
`<a:buFont typeface="+mj-lt"/>` ahead of the number and write the default
`startAt="1"` explicitly, matching PowerPoint-authored output.
