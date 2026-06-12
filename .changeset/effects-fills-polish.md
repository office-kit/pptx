---
'pptx-kit': minor
'pptx-kit-preview': minor
---

feat: effects & fills polish. The reflection effect (`a:reflection`) now
renders as a vertically mirrored, gradient-masked copy honoring start/end
alpha, distance, and the signed `sy` scale; picture bullets (`a:buBlip`)
render as real inline images in both text layout modes via the new core
reader `getParagraphBulletImageBytes` (the "■" fallback remains only when
bullet bytes are genuinely unavailable); and gradient fills inherited
through the placeholder layout/master cascade resolve via the new
`getShapeGradientFillEffective` instead of painting a hardcoded orange
tint.
