---
'@office-kit/pptx': patch
---

Numbered lists now count 1. 2. 3. in LibreOffice too

`bullets: 'number'` wrote `startAt="1"` on every paragraph. LibreOffice reads each one as a fresh start and numbers the list 1. 1. 1. The attribute is now left out; 1 is its default, so the reference desktop app numbers the list the same as before.
