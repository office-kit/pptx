---
'pptx-kit-preview': minor
---

Sharpen preview-renderer fidelity across the sample corpus. Block arrows now use
the correct OOXML shaft/head proportions (head length scales with `min(w, h)`),
text in non-rectangular autoshapes (triangle, diamond, pentagon, star, double
arrow) wraps inside the shape's inscribed text rectangle, vertical text honours
the rotated text-box insets, glow effects render as a saturated ring instead of
a pale haze, hyperlink runs take the theme `hlink` colour, bullets size to the
paragraph's first run and follow centred/right-aligned text, line breaking is
space-inclusive (matching LibreOffice/PowerPoint), the first text baseline gets
the same leading drop for every anchor, and category line charts plot at band
centres with title/axis text sized in pixels. Overall mean fg-SSIM rises from
≈0.82 to ≈0.87.
