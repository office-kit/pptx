---
'pptx-kit': minor
---

feat(site/playground): show `validatePresentation` results. The
playground now runs the validator after parsing and surfaces any
issues in a dedicated panel (with severity tint and the offending
part name when available). Lets users spot missing rels, dangling
slide IDs, etc. without dropping into the test harness.
