---
'@office-kit/pptx': minor
---

Rename `setShapeBullets` to `setShapeBulletStyle`. The old plural name read like it set the list's *content*, so calls meant to create a bulleted list reached for it and silently got back an unchanged, text-less shape. The new name says what the function actually does: it restyles the bullet glyph on paragraphs that already exist. Bulleted text is still authored in one call — `setShapeText(shape, 'A\nB\nC', { bullets: 'bullet' })`.
