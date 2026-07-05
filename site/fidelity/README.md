# Preview fidelity harness

Measures how close @office-kit/pptx's preview renderer (the `@office-kit/pptx-preview`
package) is to a real presentation engine, **as a number per slide** — the
measurement spine from the preview-fidelity roadmap. Without it, "make the
preview perfect" is unfalsifiable eyeballing.

```
ground truth (LibreOffice / PowerPoint)           →  PDF  →  PPM  ┐
                                                                  ├─→  SSIM + diff  →  report
@office-kit/pptx-preview/node · renderSlideToRgba(width)  →  SVG → resvg → RGBA  ┘
```

`renderSlideToRgba` (from `@office-kit/pptx-preview/node`) lays text out as pure SVG
`<text>` (no `<foreignObject>`) using a fontkit measurer, so resvg can
rasterize it without a browser. The measurer, resvg's `fontFiles`, and
LibreOffice all use the same bundled substitute fonts (in the package's
`fonts/`), so the engine's wrap/position math agrees with the painted pixels.

## Run it

```bash
pnpm --filter @office-kit/pptx-site fidelity                 # all samples, LibreOffice
pnpm --filter @office-kit/pptx-site fidelity -- --ours-only  # render only, no ground truth
pnpm --filter @office-kit/pptx-site fidelity -- ../samples/out/10-tables.pptx   # one file
GROUND_TRUTH=powerpoint pnpm --filter @office-kit/pptx-site fidelity            # local PP check (macOS)
```

Output lands in `site/fidelity/out/` (git-ignored): `index.html` (side-by-side
ground truth · @office-kit/pptx · diff, colored by fg-SSIM), `results.json`, and the
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
  engine (`@office-kit/pptx-preview`'s `packages/preview/src/text-layout.ts`) are
  unit-tested in `test/fidelity-metric.test.ts` / `test/text-layout.test.ts`
  and run in CI without any external renderer.

## Vertical metrics

LibreOffice / GDI place the baseline at the font's `usWinAscent` and size the
line box as `usWinAscent + usWinDescent` (no extra gap) unless the font sets the
OS/2 `USE_TYPO_METRICS` bit, in which case the `sTypo*` metrics + `typoLineGap`
apply. The measurer (`@office-kit/pptx-preview`'s `measure.ts`) mirrors this; using
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
pnpm --filter @office-kit/pptx-site fidelity -- --record

# Gate CI (or a local branch) against the committed baseline:
pnpm --filter @office-kit/pptx-site fidelity -- --check
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

## Current state (LibreOffice 24.2, 23 samples / 35 slides, 1280px)

Overall mean fg-SSIM ≈ **0.78** (plain SSIM ≈ 0.97); excluding the two
documented table-divergence slides below, ≈ **0.81**. The big lift over the
previous ≈0.66 came from an fg-SSIM-vs-shift offset search across the corpus:
it showed the entire text layer painting exactly 1px right of LibreOffice's
raster on every sample (now compensated by the `GRID_NUDGE_X` calibration in
`text-layout.ts`), master-`bodyStyle` bullets not being inherited by body
placeholders (now resolved through the paragraph cascade), chart legends being
invented for charts that author no `<c:legend>`, and the value axis missing
Excel-style headroom above the data max. Text-heavy slides now score
0.87–0.95.

Documented divergences (scored against their committed baseline, not fixed):

- `10-tables` ≈0.15 and `21-showcase` slide 4 ≈0.18 — we paint PowerPoint's
  built-in table-style header/banding fills for a `tableStyleId` whose
  definition isn't in the package; LibreOffice ships no built-in styles and
  renders plain. PowerPoint behavior wins per the project rules.
- `22-vertical-text` ≈0.33 / `23-columns` ≈0.32 — the layouts are correct in
  shape (rotation, stacking direction, sequential column fill) but LibreOffice
  auto-grows these text boxes and re-wraps to the grown extent, which shifts
  every line's break points.
- `11-charts` slide 1 ≈0.55 — remaining chart-chrome differences (tick-mark
  glyphs, marker shape, plot-area proportions).

Exact numbers move with the LibreOffice version; see the Baseline gate section
above for how to update `baseline.json` when they shift.
