---
'pptx-kit': minor
---

feat: `getParagraphBulletStyle(pres, shape, p)` returns the
paragraph-level bullet overrides — color (theme-resolved), percent
size, fixed-point size, font face — from `<a:buClr>` / `<a:buSzPct>`
/ `<a:buSzPts>` / `<a:buFont>`. Playground projects each onto the
bullet `<span>`, so decks that style bullets in an accent color or
sized-up font (a common branding move) render correctly instead of
falling back to the body's color.
