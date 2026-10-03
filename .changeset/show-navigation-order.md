---
'@office-kit/pptx-dev': patch
'@office-kit/pptx-preview': patch
---

Resolve next, previous, first, and last slide links against the active slide show's order, including repeated slides in custom shows. SVG navigation links now retain their action as a `#pptx-*` fragment so playback can choose the correct destination.

Keep presenter playback active when it is opened immediately after leaving a fullscreen presentation.
