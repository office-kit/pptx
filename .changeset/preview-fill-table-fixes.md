---
'pptx-kit-site': patch
---

site(preview): fix two fidelity bugs surfaced by the harness — (1) shapes with
an unresolved `inherit` fill (content placeholders, text boxes, bare autoshapes)
now render transparent instead of a spurious light-grey box that obscured real
content; (2) table cell text renders at the 18pt cell default, top-anchored
(ECMA default), instead of a tiny fixed 10px.
