---
'pptx-kit': minor
---

feat: `getShapeImageLinkUrl(shape)` returns the external URL of a
picture whose `<a:blip>` carries an `r:link` (Insert > "Link to file")
instead of `r:embed`. Bytes for these aren't in the package; the
playground now shows the linked URL in the placeholder rather than a
generic "no bytes" label.
