---
'@office-kit/pptx-dsl': minor
---

feat: rich runs and merges in `Table` cells

`styleCell` and the cell styles (0.4.0) cover a cell's fill, format, alignment, anchor and borders. This adds the two things a style cannot express.

- A cell in `rows` is a string or a cell object, `{ text }` or `{ paragraphs }` with optional `colSpan` / `rowSpan`. Strings behave as before and can be mixed with objects; an empty cell is the string `''`. The new exported types are `TableCell` and `TableCellSpec`.
- `paragraphs` takes core `ParagraphSpec[]`, for bold or colored runs inside one cell. The merged cell style (`cellStyle`, `headerStyle`, `styleCell`) is the base of every run, so a run states only what differs; a `Raw` callback used to lose the table-wide size and color. The style's `align` applies unless a paragraph has its own. `styleCell` receives a rich cell's `value` as its run texts joined, one line per paragraph.
- A merge is declared on its top-left cell with `colSpan` / `rowSpan`. `rows` stays a full rectangular grid and every covered position is written as `''`; any other content there is an error. Covered positions get no style and no `styleCell` call, because PowerPoint paints a merged block from its top-left cell alone.
