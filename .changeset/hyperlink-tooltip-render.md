---
'pptx-kit': minor
---

feat(site/playground): hyperlink tooltips. Shape and per-run
hyperlinks now surface their `<a:hlinkClick tooltip="…"/>` text —
shapes get an SVG `<title>` child on the `<a>` wrapper, runs get a
`title=` attribute on the HTML anchor. PowerPoint shows these on
hover during the slideshow; the playground now does too.
