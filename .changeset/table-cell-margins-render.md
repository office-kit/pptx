---
'pptx-kit': minor
---

feat(site/playground): table cell text honors authored `<a:tcPr
marL/marR/marT/marB>` insets. The renderer previously hard-coded a
4-pixel pad on every side; it now converts each EMU-valued margin to
px (falling back to 4px only when the side isn't authored) so cells
with custom inner padding line up the way PowerPoint shows them.
