---
"@office-kit/pptx": minor
"@office-kit/pptx-dev": minor
---

Say how a clip plays: `getShapeMediaPlayback` and `setShapeMediaPlayback`.

A deck could embed a video or a sound but not state anything about playing it,
so every clip waited for a click at PowerPoint's default volume. The new pair
reads and writes autoplay, loop, volume, mute, hide-when-stopped and (video
only) full screen, from the clip's media time node. Omitted properties keep
their current value.

Trimming a clip is still not supported: PowerPoint stores it in a 2010
extension rather than in the core schema.

The editor exposes the settings through its properties panel and command
palette, in English and Japanese.
