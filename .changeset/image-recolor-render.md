---
'pptx-kit': minor
---

feat(site/playground): grayscale + biLevel image filters in the
playground. The filter pipeline now composes:

1. brightness + contrast (linear feComponentTransfer)
2. grayscale (luminance-preserving feColorMatrix) when
   `<a:blip><a:grayscl/>` is set
3. biLevel two-tone (discrete tableValues snapped at the authored
   threshold) when `<a:blip><a:biLevel thresh="…"/>` is set

Pictures with Color > Grayscale or Color > Black and White now
render with the same visual treatment PowerPoint shows.
