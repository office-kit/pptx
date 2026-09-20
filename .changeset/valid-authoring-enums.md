---
"@office-kit/pptx": patch
---

Reject invalid enum values in shape, text, table, chart, and slide authoring APIs with descriptive errors instead of writing invalid PowerPoint XML or silently selecting another mode. Valid ECMA-376 shape presets remain supported, including presets outside the TypeScript autocomplete list.
