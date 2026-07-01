---
'pptx-kit': minor
---

`setShapeRunFormat` / `getShapeRunFormat` / `getShapeRunFormatEffective` now
support `fontEastAsian`, a per-run East Asian typeface override (`<a:ea>`),
alongside the existing `font` (`<a:latin>`). Previously a run's CJK glyphs
always fell back to whichever East Asian font the theme's major/minor font
scheme happened to carry, with no way to set a distinct typeface (e.g. a
serif headline vs. a sans-serif body) on an individual run. `fontEastAsian`
resolves theme `+mj-ea`/`+mn-ea` tokens and falls back to the theme's major/
minor East Asian font the same way `font` already does for Latin text.
