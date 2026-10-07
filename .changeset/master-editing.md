---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': minor
---

Slide, notes and handout masters can now be edited:

- Slide masters: `addSlideMaster`, `removeSlideMaster`, `getSlideMasterLayouts`, `getSlideMasterPlaceholders`, `getSlideMasterName` / `setSlideMasterName`, `isSlideMasterPreserved` / `setSlideMasterPreserved` and `setSlideMasterPlaceholderIncluded`. A new master gets PowerPoint's default structure: eleven Office layouts and its own copy of the theme.
- Layouts: `addSlideLayout`, `removeSlideLayout`, `addSlideLayoutPlaceholder` (content, text, picture, chart, table, SmartArt, media and online image, plus vertical content and text), `setSlideLayoutTitleIncluded`, `setSlideLayoutFootersIncluded` and `setSlideLayoutBackgroundGraphicsHidden`.
- Removing a master or layout that slides still use throws, as does removing the last one.
- Notes and handout masters: `getNotesMasterPlaceholders` / `setNotesMasterPlaceholderIncluded`, `getHandoutMasterPlaceholders` / `setHandoutMasterPlaceholderIncluded`, `getNotesPageSize` / `setNotesPageOrientation` and `getHandoutSlidesPerPage` / `setHandoutSlidesPerPage`. If the deck has no such master, the first edit creates PowerPoint's default one.
- `@office-kit/pptx-preview` adds `renderSlideLayoutToSvg`, which draws a layout's or master's background and decorative shapes without its placeholders.
- In the editor, the Slide Master, Handout Master and Notes Master tabs no longer have disabled commands:
  - Insert Slide Master, Insert Layout, Delete, Rename, Preserve, Master Layout and Insert Placeholder all work, as do the Title, Footers and Hide Background Graphics checkboxes.
  - Orientation, the placeholder checkboxes and Slides Per Page also work.
  - The master views draw the deck's decorative shapes and its real placeholder positions.
