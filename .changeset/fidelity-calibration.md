---
'pptx-kit': minor
'pptx-kit-preview': minor
---

feat: fidelity calibration sweep — measured against LibreOffice ground truth,
mean fg-SSIM rose from ≈0.66 to ≈0.78 (≈0.81 excluding documented
divergences). Body placeholders now inherit the master `bodyStyle` bullet and
hanging indent through the paragraph cascade (new `bullet` field on
`ParagraphProperties` from `getParagraphPropertiesEffective`); charts no
longer invent a legend when the XML authors no `<c:legend>`, and the value
axis gets Excel-style headroom above the data max with the tick step
preserved; the chart builder writes `<c:smooth val="0"/>` explicitly on line
series (the schema default for an absent element is smooth=1, which made
LibreOffice draw unauthored lines as curves); and the pure-SVG text layer is
nudged 0.75px left to land on LibreOffice's pixel grid.
