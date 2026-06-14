---
'pptx-kit': patch
---

Fix unstyled, broken-looking tables from `addSlideTable`

`addSlideTable` set the `firstRow` / `bandRow` flags but never wrote a
`<a:tableStyleId>`, and `createPresentation` shipped no `tableStyles.xml` part.
With no style to resolve against, PowerPoint painted the table as a borderless,
unstyled block — a "broken" grid with no rules.

- **Tables now reference PowerPoint's "No Style, Table Grid" built-in**
  (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`) via `<a:tableStyleId>`, the same
  default PptxGenJS and PowerPoint itself emit, so a table resolves to a clean
  ruled grid. Callers can override with the internal `styleId` option (or the
  existing `setTableStyleId`).
- **`createPresentation` now ships `/ppt/tableStyles.xml`** (referenced from
  `presentation.xml.rels`), matching every PowerPoint-authored deck, so the
  `tableStyleId` always has a backing part.
