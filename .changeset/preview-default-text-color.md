---
'pptx-kit': minor
'pptx-kit-preview': patch
---

Preview: take the default text color from the deck's body style, not the `tx1` token

The preview used `scheme:tx1` as the fallback color for runs without an authored color. On a template with an inverted color map (`tx1 → lt1`) that resolves to the light slot, so body text was painted white on the white background — the whole slide looked blank. PowerPoint instead takes the fallback from the master `bodyStyle` (e.g. `schemeClr bg1`). The preview now does the same via the newly exported `resolveDeckBodyTextColor(slide)`, so default-colored text and table-cell text resolve to the color PowerPoint actually paints.

- New export **`resolveDeckBodyTextColor(slide)`** — the deck's resolved body-text color (master `bodyStyle`, run through the effective color map + theme). This is the color `addSlideTable` / `addSlideChart` bake in, now reusable by renderers.
