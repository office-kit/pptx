---
'@office-kit/pptx': patch
'@office-kit/pptx-dev': patch
---

Image brightness and contrast now work on image-filled shapes. The editor preserves these corrections when switching fill types and restoring the image.

Image corrections, opacity, and DrawingML color transforms now read both fixed-point and percent-suffixed values correctly. Hue offsets use angle units, preventing incorrect colors in imported presentations.
