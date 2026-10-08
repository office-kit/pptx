---
'@office-kit/pptx-editor': minor
'@office-kit/pptx-dev': minor
'@office-kit/pptx-preview': minor
---

feat: copy slides and objects to the system clipboard. Copying slides or objects in the editor now puts them on the clipboard for other presentation apps and documents: slides and drawn objects paste there as pictures at their size on the slide, and text boxes and tables copied on their own paste as text and tables. Another editor — in another tab or browser — pastes them back as editable slides and objects, with their pictures, layouts and themes. Pictures and text copied in other apps paste as a picture or a text box.

feat: the slide thumbnail menu has **Download Selected Slides...**, which saves just the selected slides as a `.pptx` to import into another presentation app.

feat: `renderSlideToSvg(pres, slide, { background: false })` draws only the slide's own shapes on a transparent surface, without the background or master and layout graphics.
