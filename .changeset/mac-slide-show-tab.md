---
'@office-kit/pptx': minor
'@office-kit/pptx-dev': minor
---

`SlideShowProperties` gains an optional `showMediaControls` (PowerPoint's Show Media Controls, stored as `p14:showMediaCtrls`); `getSlideShowProperties` always reports it. The editor's Slide Show tab adds Rehearse Timings (time each slide while presenting, then keep the times as slide timings), Record, and the Use Timings, Play Narrations and Show Media Controls options; the Microsoft 365-only commands are shown disabled.
