---
"@office-kit/pptx": patch
---

Replace text across adjacent formatting runs in slide text, including table cells. Replacement text inherits the first matched run’s format, while surrounding text and run metadata stay intact. Matches stop at paragraph and explicit line-break boundaries. Regular expressions now operate on each complete text segment instead of each individual run, so anchors and replacement context follow the visible segment text. The return value remains the number of changed text nodes.
