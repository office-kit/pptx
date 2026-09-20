---
'@office-kit/pptx': patch
---

fix: theme readers and the preview picked the first theme part by name instead of the slide master's theme

`getPresentationTheme`, `getPresentationFonts`, `setPresentationTheme`, `setPresentationFonts` and every color resolver built on them (slide backgrounds, `scheme:*` fills, the deck body text color, the preview renderer) read `/ppt/theme/theme1.xml` whenever a package carried more than one theme. Notes and handout masters carry their own theme, and Google Slides exports write the notes theme as `theme1.xml` and the slide theme as `theme2.xml`, so a dark deck rendered on white with black text.

The deck theme is now the one the first slide master (in `<p:sldMasterIdLst>` order) relates to, then the presentation part's own theme relationship, and only then the first theme part by name.
