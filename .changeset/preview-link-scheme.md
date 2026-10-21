---
'@office-kit/pptx-preview': patch
'@office-kit/pptx-editor': patch
---

The preview no longer turns a deck's hyperlinks into `<a href>` unless they point to `http:`, `https:`, `mailto:` or a place inside the deck. A `javascript:`, `data:` or `vbscript:` link (in any letter case, or with spaces, tabs or line breaks around or inside the scheme) in a run, a table cell or a shape click action used to be written as a live link, so an SVG from an untrusted deck could run script in the page that shows it; the text now renders without the link. Font names from the theme and from runs are now quoted in inline styles, so a name containing `;` can no longer add CSS declarations such as `background-image:url(…)`. The editor is republished to require this preview.
