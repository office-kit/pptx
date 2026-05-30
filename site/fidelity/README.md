# Preview fidelity harness

Measures how close pptx-kit's preview renderer (`renderSlideSvg`) is to a real
presentation engine, **as a number per slide** — the measurement spine from
[Phase 0 of the preview-fidelity roadmap](../../). Without it, "make the
preview perfect" is unfalsifiable eyeballing.

```
ground truth (LibreOffice / PowerPoint)   →  PDF  →  PPM  ┐
                                                          ├─→  SSIM + diff  →  report
pptx-kit renderSlideSvg  →  SVG  →  resvg  →  RGBA  ──────┘
```

## Run it

```bash
pnpm --filter pptx-kit-site fidelity                 # all samples, LibreOffice
pnpm --filter pptx-kit-site fidelity -- --ours-only  # render only, no ground truth
pnpm --filter pptx-kit-site fidelity -- ../samples/out/10-tables.pptx   # one file
GROUND_TRUTH=powerpoint pnpm --filter pptx-kit-site fidelity            # local PP check (macOS)
```

Output lands in `site/fidelity/out/` (git-ignored): `index.html` (side-by-side
ground truth · pptx-kit · diff, colored by SSIM), `results.json`, and the
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

- **SSIM** ∈ [-1, 1], 1.0 = identical. It is computed on luma over an 8×8
  sliding window (Wang et al., 2004).
- **Baselines read high (~0.95) even with obvious gaps**, because slides are
  mostly white and SSIM rewards matching flat regions. Treat SSIM as a
  _relative_ ranking across slides/commits, not an absolute "% correct". The
  diff image is the ground truth for _what_ is wrong; SSIM is for _which slides
  regressed_. (Foreground-weighted scoring is a future refinement.)
- The metric core (`image.ts`, `ssim.ts`, `ppm.ts`, `png.ts`) is unit-tested in
  `test/fidelity-metric.test.ts` and runs in CI without any external renderer.

## Known limitation this harness is built to expose

The current renderer lays text out in SVG `<foreignObject>`, which **resvg (any
browser-free rasterizer) cannot render**. So text is absent from our rasters
today, and it dominates the diff. That is the point: the gap quantifies the
"move text to self-laid-out pure SVG" work the roadmap schedules for Phase 1,
which is also what unblocks browser-free Node rendering.

## First baseline (LibreOffice 26.2, 21 samples, 1280px)

Overall mean SSIM ≈ **0.958**. Weakest first — i.e. where to look next:
`10-tables` ≈ 0.87, `07-fills` ≈ 0.92, `11-charts` ≈ 0.94, `03-text-formatting`
≈ 0.94. Exact numbers move with the LibreOffice version, so no baseline file is
committed yet; CI will establish its own once the gate lands (Phase 4).
