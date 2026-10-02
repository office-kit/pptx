---
"@office-kit/pptx-dev": patch
---

Fix extending outline text selections upward across slides with Shift+Up. Keep the original selection anchor when selecting backwards, including consecutive key presses, so copying and editing apply to the whole selected range.
