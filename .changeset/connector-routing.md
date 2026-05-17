---
'pptx-kit': minor
---

feat(site/playground): bent / curved connector routing.

`bentConnector{2,3,4,5}` render as the matching L / Z / two-step /
three-step paths, and `curvedConnector{2,3,4,5}` render as quadratic
/ cubic Bézier curves between the connector's bounding-box endpoints.
Previously every connector preset projected to a straight line; flow-
chart and diagram decks now show the right cadence.
