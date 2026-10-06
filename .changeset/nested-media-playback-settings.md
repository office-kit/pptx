---
'@office-kit/pptx': patch
'@office-kit/pptx-dev': patch
---

Show and edit playback settings for audio and video nested inside animation timing groups. Preserve enclosing start conditions when changing volume, looping, or video display settings, and account for parent delays when reading and editing simple automatic playback. Reject unsupported start-condition edits before changing the document. Also reject non-finite volume values instead of writing invalid XML.
