---
"@office-kit/pptx-dev": patch
---

Preserve exact trim and fade positions when opening media whose duration is not a multiple of 50 milliseconds. Keep keyboard and drag adjustments in 50 millisecond increments without rounding saved values or losing the clip endpoint.
