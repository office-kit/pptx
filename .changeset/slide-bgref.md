---
'pptx-kit': minor
---

feat: `getSlideBackground` now handles `<p:bgRef>` (the theme-fill-
reference variant of slide background, e.g. `<p:bgRef idx="1003">
<a:schemeClr val="bg1"/></p:bgRef>`). Returns the inner color element
as a solid fill so renderers paint the slide background even when
the deck uses the theme-reference form instead of explicit `<p:bgPr>
<a:solidFill>`.
