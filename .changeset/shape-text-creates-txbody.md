---
"@office-kit/pptx": patch
---

fix: setShapeText / appendShapeText now add text to a shape that has no text body

Previously, setting text on a shape authored without one (e.g. `addSlideShape`
called without `text`) threw `shape "…" has no <p:txBody>`. PowerPoint always
gives an autoshape a text body so you can click in and type, so these functions
now create the body on demand and populate it, matching that behavior. Picture /
table shapes still throw, since they are not text-bearing.
