---
'@office-kit/pptx-dev': minor
---

The editor's right-click menus for text, table cells, pictures and the slide background now match the reference desktop app (Mac):

- **Text being edited** gets the reference desktop app's menu instead of the browser's: Cut, Copy, Paste, Exit Edit Text, Font..., Paragraph..., Bullets ▸ and Numbering ▸ galleries, Format Shape..., Lock, Hyperlink... and New Comment. The menu leaves the caret and selection in the text, and its commands act on the selected range.
- **Table cells** offer Insert ▸ (columns left/right, rows above/below), Delete ▸ (columns, rows, table), Select ▸ (table, column, row), Merge Cells and Split Cells..., plus the text commands. Clear cell text, Select all cells and Select table left the menu (Delete, ⌘A and Select ▸ Select Table do the same).
- **Pictures** offer Change Picture ▸ From a File..., Crop, Format Picture... and the reference desktop app's other picture items.
- **The slide background** offers Paste Special..., New/Duplicate/Delete Slide, Hide Slide, Ruler, Grid and Guides ▸ toggles, Zoom..., Format Background..., Slide Show and New Comment; Layout ▸, Reset Slide and Select all left it (they remain on the Home tab and ⌘A).

Menus show the reference desktop app's Mac shortcut hints, including in submenus; commands the editor cannot perform yet are disabled with the reason as a tooltip.
