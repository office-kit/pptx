# Preview fidelity harness

Measures how close pptx-kit's preview renderer (the `@pptx-kit/preview`
package) is to a real presentation engine, **as a number per slide** — the
measurement spine from the preview-fidelity roadmap. Without it, "make the
preview perfect" is unfalsifiable eyeballing.

```
ground truth (LibreOffice / PowerPoint)           →  PDF  →  PPM  ┐
                                                                  ├─→  SSIM + diff  →  report
@pptx-kit/preview/node · renderSlideToRgba(width)  →  SVG → resvg → RGBA  ┘
```

`renderSlideToRgba` (from `@pptx-kit/preview/node`) lays text out as pure SVG
`<text>` (no `<foreignObject>`) using a fontkit measurer, so resvg can
rasterize it without a browser. The measurer, resvg's `fontFiles`, and
LibreOffice all use the same bundled substitute fonts (in the package's
`fonts/`), so the engine's wrap/position math agrees with the painted pixels.

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
  engine (`@pptx-kit/preview`'s `packages/preview/src/text-layout.ts`) are
  unit-tested in `test/fidelity-metric.test.ts` / `test/text-layout.test.ts`
  and run in CI without any external renderer.

## Vertical metrics

LibreOffice / GDI place the baseline at the font's `usWinAscent` and size the
line box as `usWinAscent + usWinDescent` (no extra gap) unless the font sets the
OS/2 `USE_TYPO_METRICS` bit, in which case the `sTypo*` metrics + `typoLineGap`
apply. The measurer (`@pptx-kit/preview`'s `measure.ts`) mirrors this; using
fontkit's hhea `.ascent` instead misplaces the baseline by ~12px at 44pt.

**Center / bottom anchoring** carries one extra wrinkle: LibreOffice/PowerPoint
sit a vertically-centered (or bottom-anchored) line slightly _lower_ than the
win-metric line box predicts — measured against ground truth (an IoU offset
search) as ≈0.036 of the line's (ascent+descent), independent of font size and
isolated to non-top anchoring. `text-layout.ts` applies that as the documented
`CENTER_ANCHOR_DROP` constant; top-anchored text needs no correction and gets
none. (No clean font-metric formula reproduces it, so it is an empirically
calibrated constant, validated to lift every affected slide with no regression.)

## Baseline gate

The fidelity gate commits a `baseline.json` snapshot of per-slide fg-SSIM
scores so CI can detect regressions automatically.

### How it works

```bash
# After a renderer change, record a fresh baseline:
pnpm --filter pptx-kit-site fidelity -- --record

# Gate CI (or a local branch) against the committed baseline:
pnpm --filter pptx-kit-site fidelity -- --check
```

`--record` writes `site/fidelity/baseline.json` (sorted keys, trailing newline)
and prints the path. `--check` compares the current run against that file and
exits 1 if:

- Any slide's fg-SSIM drops more than **0.02** below its baseline value.
- A sample file or slide appears in the run but has no baseline entry.
- A baseline entry has no corresponding sample in the run.

The **0.02 tolerance** matches one full LibreOffice point-release of rendering
jitter (~0.01 per release) with headroom, so version bumps on the runner don't
produce phantom failures.

`--check` also prints any slides that improved by more than 0.02 with a hint to
re-record, so the baseline tracks genuine progress rather than just blocking
regressions.

`--check` and `--record` are mutually exclusive with `--ours-only`.

### CI bootstrap and update procedure

Every `--check` run writes a `site/fidelity/baseline.candidate.json` file (same
shape as `baseline.json`) so the CI job always uploads it as the
`fidelity-baseline-candidate` artifact regardless of pass/fail.

**First time / after a renderer improvement:**

1. Download the `fidelity-baseline-candidate` artifact from the CI run.
2. Copy the file to `site/fidelity/baseline.json`.
3. Commit and push. The gate will pass on the next run.

**After a LibreOffice runner upgrade that shifts absolute scores:**

Follow the same procedure — download the candidate, inspect the diffs, commit
if the changes are expected.

## Current state (LibreOffice 26.2, 21 samples, 1280px)

Overall mean fg-SSIM ≈ **0.66** (plain SSIM ≈ 0.97). Text is horizontally
near-exact, correctly sized, and now vertically aligned for both top- and
center/bottom-anchored blocks (e.g. `01-title-only` 0.26→0.65, `20` 0.28→0.72,
`15`/`16` ~0.28→~0.75 after the center-anchor calibration). Graphic/image slides
score high (`09-images` 0.98, `14-background` 0.97, `13-zorder` 0.95). Remaining
laggards: `10-tables` (≈0.14 — we paint PowerPoint's built-in table-style
header/banding fills, which LibreOffice renders plain: a documented
LO-vs-PowerPoint divergence) and `05-preset-shapes` (≈0.38 — tiny centered
labels). A residual ~1px horizontal text offset (resvg-vs-pdftoppm pixel grid)
caps per-glyph overlap. Bullets/vertical/column text and per-run table-cell
styling are partially or not yet ported. Exact numbers move with the LibreOffice
version; see the Baseline gate section above for how to update `baseline.json`
when they shift.
