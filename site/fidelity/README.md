# Preview fidelity harness

Measures how close pptx-kit's preview renderer (`renderSlideSvg`) is to a real
presentation engine, **as a number per slide** — the measurement spine from
the preview-fidelity roadmap. Without it, "make the preview perfect" is
unfalsifiable eyeballing.

```
ground truth (LibreOffice / PowerPoint)        →  PDF  →  PPM  ┐
                                                               ├─→  SSIM + diff  →  report
renderSlideSvg(textLayout:'svg', measureText)  →  SVG  →  resvg  →  RGBA  ┘
```

Text is laid out as pure SVG `<text>` (no `<foreignObject>`) using a fontkit
measurer, so resvg can rasterize it without a browser. The measurer, resvg's
`fontFiles`, and LibreOffice all use the same bundled substitute fonts
(`fonts/`), so the engine's wrap/position math agrees with the painted pixels.

## Run it

```bash
pnpm --filter pptx-kit-site fidelity                 # all samples, LibreOffice
pnpm --filter pptx-kit-site fidelity -- --ours-only  # render only, no ground truth
pnpm --filter pptx-kit-site fidelity -- ../samples/out/10-tables.pptx   # one file
GROUND_TRUTH=powerpoint pnpm --filter pptx-kit-site fidelity            # local PP check (macOS)
```

Output lands in `site/fidelity/out/` (git-ignored): `index.html` (side-by-side
ground truth · pptx-kit · diff, colored by fg-SSIM), `results.json`, and the
per-slide PNGs.

Flags: `--width <px>` (default 1280), `--engine libreoffice|powerpoint`,
`--out <dir>`, `--samples <dir>`, `--ours-only`.

## Prerequisites

- **LibreOffice** (the CI / default ground truth): macOS `brew install --cask
libreoffice`; Debian/CI `apt-get install libreoffice`. Override the binary
  with `PPTX_KIT_SOFFICE`.
- **poppler / `pdftoppm`** (PDF → raster): macOS `brew install poppler`;
  Debian/CI `apt-get install poppler-utils`. Override with `PPTX_KIT_PDFTOPPM`.
- **PowerPoint** (optional, local, macOS only): the high-fidelity local check.
  CI never needs it.

`--ours-only` needs none of the above — useful for a quick render smoke test.

## How to read the numbers

- **fg-SSIM** (foreground-weighted SSIM) is the headline metric. Each window is
  weighted by how much "ink" the ground truth has there, so blank slide areas —
  which dominate plain SSIM and reward a renderer for drawing _nothing_ — barely
  count. Missing or misplaced text scores low here; correctly rendered text
  scores high. Use it as the relative gauge across slides/commits.
- **plain SSIM** is reported too, but on mostly-white slides it reads
  deceptively high (~0.95) even when text is absent, because the white field
  matches. It is the weaker signal.
- fg-SSIM is **strict on text**: text is high-frequency, so even a few-pixel
  offset drops the inked-window correlation sharply. A high fg-SSIM means the
  text is genuinely pixel-aligned, not just "looks about right".
- The diff image is the ground truth for _what_ is wrong; fg-SSIM tells you
  _which_ slides to look at.
- The metric core (`image.ts`, `ssim.ts`, `ppm.ts`, `png.ts`) and the layout
  engine (`../src/lib/playground/text-layout.ts`) are unit-tested in
  `test/fidelity-metric.test.ts` / `test/text-layout.test.ts` and run in CI
  without any external renderer.

## Vertical metrics

LibreOffice / GDI place the baseline at the font's `usWinAscent` and size the
line box as `usWinAscent + usWinDescent` (no extra gap) unless the font sets the
OS/2 `USE_TYPO_METRICS` bit, in which case the `sTypo*` metrics + `typoLineGap`
apply. The measurer (`measure.ts`) mirrors this; using fontkit's hhea `.ascent`
instead misplaces the baseline by ~12px at 44pt.

## Current state (LibreOffice 26.2, 21 samples, 1280px)

Overall mean fg-SSIM ≈ **0.46** (plain SSIM ≈ 0.96). Text layout is horizontally
near-exact and correctly sized; top-anchored text bodies align well (e.g.
`03-text-formatting` ≈ 0.57), graphic-heavy slides score high (`09-images` ≈
0.95). Known laggards, tracked for follow-up PRs: centered single lines carry a
~4px vertical residual; bullets/measured-autofit, vertical/column text, and
table-cell run styling are not yet ported to the pure-SVG path; a spurious faint
placeholder fill and near-white preset shapes hurt `05`/`10`. Exact numbers move
with the LibreOffice version, so no baseline file is committed yet; CI will
establish its own once the gate lands.
