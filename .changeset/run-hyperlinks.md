---
'pptx-kit': minor
---

feat(site/playground): per-run hyperlinks. Runs carrying `<a:hlinkClick>`
now render in the theme's hyperlink color with an underline, and the
span is wrapped in an `<a href>` so the preview is clickable. Per-run
font / color / formatting overrides still apply on top — the link
styling fills the gaps the run didn't author.
