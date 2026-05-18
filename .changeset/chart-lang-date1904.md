---
'pptx-kit': minor
---

feat: `ChartSpec.language` (`<c:chartSpace><c:lang val=…/>`) and
`ChartSpec.date1904` (`<c:date1904 val=…/>`) — chartSpace-level Office
metadata round-tripped for parity. `language` is the Office UI
language code (e.g. `'en-US'`, `'ja-JP'`); `date1904` selects the
1904 date epoch (default `false` = Excel 1900 epoch, surface only
when explicitly true). pptx-kit's renderers don't act on either yet.
