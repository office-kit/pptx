---
'@office-kit/pptx-editor': minor
---

New package: embed the PowerPoint-style presentation editor in your own web application.

- `mountEditor(target, { source, locale, onSave })` opens a `.pptx` (or a new presentation) in `target` and returns a handle with `ready`, `save()` and `destroy()`. `onSave` receives the `.pptx` bytes when the user saves; without it, saving downloads the file.
- The editor renders in a shadow root: the page's styles do not reach it and its styles do not reach the page. Its keyboard shortcuts leave the page's own inputs alone, and several editors can share a page.
- The package ships compiled JavaScript with the Svelte runtime bundled, so the host application needs no framework.
