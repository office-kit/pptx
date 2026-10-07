---
'@office-kit/pptx-dev': patch
---

The editor's ruler and texture defaults now follow what Mac PowerPoint does:

- **Ruler with several paragraphs selected** shows the last selected paragraph's indent markers and tab stops, as PowerPoint does, instead of the first paragraph's.
- **Picture or texture fill** inserts the last texture picked for a shape in this session (starting with Papyrus) for both shapes and slide backgrounds. Textures picked in Format Background do not change it. Like PowerPoint's, which resets on relaunch, it is not saved and returns to Papyrus when the editor page is reloaded.
