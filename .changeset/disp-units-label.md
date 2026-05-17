---
'pptx-kit': minor
---

feat(site/playground): `<c:dispUnits>` value-axis label. When the
chart authors a display-units token (`thousands`, `millions`, etc.)
the value-axis now emits an italic "Thousands" / "Millions" /
… label rotated alongside the axis (vertical orientation) or to
the right of the rightmost tick (horizontal). Completes the
display-units rendering — values are already divided, and now the
scale self-describes.
