---
'pptx-kit': minor
---

feat: `getShapeBodyPrEffective(pres, shape)` — `<a:bodyPr>` cascade
covering anchor, wrap, vertical-text direction, and inset margins.
Walks shape → layout placeholder → master placeholder bodyPr the same
way the rPr / pPr cascades do. Playground uses it so placeholders
inherit text alignment / margins from the layout / master without
each slide having to re-author them.
