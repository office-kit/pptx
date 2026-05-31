---
'pptx-kit-site': patch
---

site(preview): render template decoration. Non-placeholder shapes defined on
the slide layout and master — logos, corner bars, divider lines, watermark
text — now appear in the preview (master behind layout, both behind the
slide's own content), rendered through the full shape pipeline. Previously
layout shapes were drawn as crude solid rectangles and master shapes / picture
logos were dropped entirely.
