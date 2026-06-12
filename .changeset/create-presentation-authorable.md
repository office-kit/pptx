---
'pptx-kit': minor
---

`createPresentation()` now returns an immediately-authorable deck

Previously `createPresentation()` returned an OPC package with only the OPC
defaults — no slide master, layouts, theme, or slide size — so
`getSlideLayouts()` came back empty and `addSlide({ layout })` was
impossible. From-scratch authoring (a headline feature in the README) did
not actually work without loading a template file.

`createPresentation()` now ships a slide master, the Office theme, and three
layouts — `Blank`, `Title Slide`, and `Title and Content` — so you can go
straight to `addSlide` / `addTitleSlide` / `addContentSlide` and `savePresentation`.
Every emitted part is validated against the ECMA-376 XSDs in CI. The slide
size defaults to 16:9 and is selectable: `createPresentation({ size: '4:3' })`.

Also in this release (input-validation hardening at the authoring boundary):

- `addSlideChart` now rejects a series `color` (and `pointColors` /
  `trendline.color` / plot- and chart-area fills / axis & gridline colors)
  that isn't an sRGB hex (`#RRGGBB` or `RRGGBB`) with a clear error, instead
  of silently emitting an invalid `<a:srgbClr val="…"/>` that PowerPoint
  dropped or repaired. Bare `RRGGBB` (no `#`) is accepted and normalized;
  scheme tokens like `accent1` are correctly rejected, since charts emit
  `srgbClr`.
- `addSlideTable` with empty `rows: []` (or a row with no cells) now throws
  an actionable `addSlideTable: …` error at the boundary rather than
  producing a grid-less `<a:tbl>` that triggers PowerPoint's repair dialog.
  (The error message previously named the old internal `addTable` path.)
- `findSlideLayout`'s case-sensitive, locale-dependent name matching is now
  documented in its JSDoc and the README, pointing readers to the
  locale-stable `findSlideLayoutByType` and to `RegExp`/`i` for
  case-insensitive name lookups. No behavior change.

No breaking changes. `createPresentation()` keeps its zero-argument call
signature; the new `{ size }` options object is optional.
