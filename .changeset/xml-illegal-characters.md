---
'@office-kit/pptx': patch
---

Reject U+FFFE, U+FFFF, and unpaired surrogates in authored XML text and attributes, preventing malformed PPTX parts and silent replacement with U+FFFD. Valid supplementary characters, including emoji, remain unchanged.
