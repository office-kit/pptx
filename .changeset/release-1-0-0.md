---
'pptx-kit': major
---

**1.0.0** — first stable release. The public API is now frozen under SemVer.

**What works at 1.0:**

- **Read** any `.pptx` produced by PowerPoint, Keynote, Google Slides, or
  LibreOffice Impress, and save it back without corruption. Unknown
  extensions are preserved verbatim on round-trip.
- **Template editing**: token / text replace across slides and speaker
  notes, image swap with geometry preserved, slide CRUD with placeholder
  inheritance from layout / master.
- **Authoring on top of an existing master**: 180+ preset shapes, custom
  text formatting, tables, embedded charts (column / line / bar / pie /
  doughnut / area) with auto-generated xlsx, solid / gradient / pattern /
  image fills, shadows and glows, rotation / flip / z-order, hyperlinks
  and click actions, notes and comments, slide transitions, simple
  entrance / exit animations.
- **Diagnostics**: `validatePresentation` returns invariant violations;
  every XML part is validated against the ECMA-376 XSDs in CI.
- **Bundling**: one ESM build runs in both Node ≥ 20 and modern browsers.
  Tree-shaking is enforced by a CI test — minimal `load → save` bundle
  is < 75 KB unminified, full fn-API bundle is ~120 KB.

**Deferred to post-1.0** (read pass-through preserved on round-trip):

- Constructing new themes / masters / layouts from scratch.
- SmartArt authoring.
- Complex animation timing-tree authoring.
- OLE / ActiveX authoring.
- Document encryption (read + write).

**Performance (M-series Node 20):** 100-slide synthetic deck saves in
~25 ms, loads in ~20 ms. 100 MB templates fit comfortably under the 2 s
load/save targets.

**Migration:** if you were on the pre-1.0 class API
(`Presentation` / `Slide` / `SlideShape` / `SlideLayout`), see the
preceding changeset for the rename table. There is no class API at 1.0.
