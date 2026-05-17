---
'pptx-kit': minor
---

feat(site/playground): default text body to the theme's font scheme.
`<a:fontScheme><a:majorFont>` becomes the default face for title /
ctrTitle placeholders; `<a:minorFont>` covers everything else. The
existing per-run `<a:rPr typeface>` override still wins. Templates
that brand-themselves to Aptos / Inter / etc. now render with their
authored fonts instead of always falling back to Calibri.
