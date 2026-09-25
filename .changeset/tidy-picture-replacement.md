---
"@office-kit/pptx": patch
---

Replacing a copied picture now changes only that picture, preserving other pictures that share its media or slide relationship. Repeated same-format replacements reuse the detached media part. Invalid image crop input now preserves the previous crop, including when a subsequent edit is saved.
