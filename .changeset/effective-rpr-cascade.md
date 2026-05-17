---
'pptx-kit': minor
---

feat: add `getShapeRunFormatEffective(pres, shape, p, r)` — resolves a
run's character properties (font, size, color, bold, italic, underline)
through the full ECMA-376 §21.1.2.4.7 inheritance chain: run `<a:rPr>` →
`<a:endParaRPr>` → paragraph `<a:defRPr>` → text-body `<a:lstStyle>` →
matching layout placeholder → matching master placeholder → master
`<p:txStyles>` → theme `<a:fontScheme>`. Theme tokens like `+mj-lt` are
expanded to the deck's major/minor typefaces. The existing
`getShapeRunFormat` still returns the literal `<a:rPr>` only.
