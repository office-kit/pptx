---
'pptx-kit': minor
---

Added two authoring capabilities aimed at complex, dense slide layouts
(process diagrams, KPI cards, branded decks):

- `groupShapes` / `ungroupShapes` compose a selection of top-level shapes
  into a single `<p:grpSp>` (and reverse it). The group's bounds are the
  union of its members; moving or resizing the group afterwards rescales
  every member on ungroup, so a "KPI card" or diagram node built from a
  rectangle + label can be treated — and repositioned — as one unit.
- `setPresentationTheme` / `setPresentationFonts` patch a deck's color
  scheme and font scheme in place, so a from-scratch deck can be branded
  with a custom palette and typography without hand-authoring a template.
  Only the slots passed in are overwritten; every other slot keeps its
  existing value.
