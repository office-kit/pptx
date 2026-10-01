---
'@office-kit/pptx-dev': patch
---

Fall back to plain text when pasted character metadata contains unsupported underline or strike styles, font sizes, spacing, or colors. Preserve valid scheme-prefixed theme colors when copying formatted text.

Ignore HTML font sizes smaller than one point so pasting web content cannot fail while applying its text formatting.
