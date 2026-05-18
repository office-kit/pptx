---
'pptx-kit': patch
---

fix(validator): `validatePresentation` now flags duplicate
`<p:cNvPr id="N">` values inside a single slide's `<p:spTree>` as
errors. PowerPoint requires every shape's non-visual ID to be unique
within its slide; duplicates often appear after pasting shapes from
another slide without re-allocating IDs. The walk recurses into
`<p:grpSp>` so duplicates nested in groups are also caught.
