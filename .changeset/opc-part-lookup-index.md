---
'@office-kit/pptx': patch
---

Fixed: loading a `.pptx` with many parts took time quadratic in the part count, because each part's content type was found by scanning every `<Override>` in `[Content_Types].xml` and each part lookup scanned every part. A package with 30,000 parts and overrides took about 7 seconds to load; it now takes about a quarter of a second.
