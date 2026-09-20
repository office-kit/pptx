---
'@office-kit/pptx-dsl': minor
---

feat: `Line` and `Group` elements, bullets and paragraph spacing on `Text`, text alignment on `Shape`, per-cell table styles

These are the pieces a typical consulting-style slide needed `Raw` for:

- `<Line x1 y1 x2 y2 color width />` draws a straight line (rules, dividers). It takes end points instead of bounds.
- `<Group name>` groups the shapes its children create, so a card made of a panel, an accent bar and two text boxes moves and resizes as one object. Groups nest.
- `Text` gains `bullets` (`'bullet'`, `'number'`, a custom character, …) and `paragraphSpacing={{ before, after }}` in points; both apply to every paragraph.
- `Shape` gains `align` and `anchor` for its `text`, so a label can sit centred in a circle or an arrow.
- `Table` cell styles gain `align` and `borders` (per side, width in points), and the new `styleCell={({ row, column, value }) => style}` callback styles one cell at a time: a coloured status column, a bold total row. Styles merge in the order `cellStyle`, `headerStyle`, `stripeFill`, `styleCell`.

Existing decks compile to the same output.
