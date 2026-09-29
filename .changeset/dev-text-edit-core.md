---
'@office-kit/pptx-dev': patch
---

fix: a double-click text edit whose new text contains `$&`, `$'` or `` $` `` no longer fails as "affects other text or slides". The edit rules now come from `@office-kit/pptx-dsl/source-edit`.
