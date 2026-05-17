---
'pptx-kit': minor
---

feat: chart builder writes back `ChartSpec.titleStyle`. Previously the
reader picked up authored `<a:rPr sz/b/i><a:solidFill>` on chart
titles but the builder dropped any incoming style, so round-tripping
(read → save → reload) lost the title font / color. The builder now
emits `<a:rPr>` attributes and an inner `<a:solidFill><a:srgbClr/>`
when a `titleStyle` is provided. New round-trip test
(`fn-chart-readback`) covers this; total tests 801.
