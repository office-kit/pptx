---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': minor
'@office-kit/pptx-dev': minor
---

The reference desktop app's 74 built-in table styles are now fully supported.

- `setTableStyleId` accepts a built-in style's English name as well as a GUID (`setTableStyleId(table, 'Light Style 1 - Accent 2')`). Applying a built-in style writes the reference desktop app's definition of it into `ppt/tableStyles.xml` (creating the part when a deck has none), as the reference desktop app does, so Keynote, Google Slides and LibreOffice draw the same table. A deck's own definition of a GUID is kept. `addSlideTable` likewise writes the definition of its default style.
- New `BUILTIN_TABLE_STYLES` lists every built-in style (`id`, `name`, gallery `category`) in the order of the reference desktop app's Table Styles gallery.
- `getTableCellAppearanceEffective` now resolves every built-in style, not only Medium Style 2 - Accent 1 and No Style, Table Grid, and reports `fillOpacity` for translucent fills. New `getTableBackgroundEffective` resolves the table background (`a:tblPr` fill, else the style's `a:tblBg`).
- The preview draws all table style parts for built-in and custom styles, including translucent bands and the Themed Styles' background.
- The editor's Table Design ▸ Table Styles gallery offers all built-in styles, grouped as in the reference desktop app (Best Match for Document, Light, Medium, Dark), with swatches drawn by the preview renderer that follow the Table Style Options check boxes, the reference desktop app's English and Japanese names, and Clear Table.
