---
'@office-kit/pptx-editor': patch
---

Fixed: Video Format ▸ Poster Frame ▸ Reset could occasionally set a blank, fully transparent poster instead of the video's first frame (seen in Chromium on Linux). Reset now waits until the first frame has actually painted before saving it as the poster.
