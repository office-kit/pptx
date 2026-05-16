# Manual-inspection samples

This directory holds a generator script that produces one `.pptx` per major
feature area under `samples/out/`. Open the files in PowerPoint, Keynote,
Google Slides, or LibreOffice Impress to verify the output renders as
intended.

## Regenerate

```sh
pnpm samples
```

The script is implemented in `test/samples-generate.test.ts` and runs only
when `GENERATE_SAMPLES=1`, so it does not churn artifacts in the normal CI
test run.

## What each sample exercises

| File                        | Covers                                                        |
| --------------------------- | ------------------------------------------------------------- |
| `01-title-only.pptx`        | Smallest possible deck — one slide with a title placeholder.  |
| `02-layouts.pptx`           | Title, section header, and content layouts side by side.      |
| `03-text-formatting.pptx`   | Font, size, color, bold / italic / underline run formatting.  |
| `04-bullets-alignment.pptx` | Bulleted list with mixed marker styles and alignment.         |
| `05-preset-shapes.pptx`     | 12 preset geometries: rect, ellipse, arrows, stars, polygons. |
| `06-lines-arrows.pptx`      | Lines + arrowheads (triangle, oval), dashed strokes.          |
| `07-fills.pptx`             | Solid, linear-gradient, and pattern fills.                    |
| `08-shadow-glow.pptx`       | Drop shadow and outer-glow effects.                           |
| `09-images.pptx`            | Three embedded PNGs in distinct colors.                       |
| `10-tables.pptx`            | 5×4 table with header and banded rows.                        |
| `11-charts.pptx`            | Column + line + pie charts in one deck.                       |
| `12-geometry.pptx`          | Shape rotation, horizontal flip, vertical flip.               |
| `13-zorder.pptx`            | Two overlapping shapes with explicit z-order reorder.         |
| `14-background.pptx`        | Solid-color slide background.                                 |
| `15-notes-comments.pptx`    | Speaker notes plus two review comments from two authors.      |
| `16-transitions.pptx`       | Slide transitions: fade, push, wipe.                          |
| `17-animations.pptx`        | Entrance / exit animations: `fadeIn`, `appear`, `fadeOut`.    |
| `18-hyperlinks.pptx`        | External URL hyperlink on a text run.                         |
| `19-token-fill.pptx`        | `{{token}}` template-fill via `replaceTokensInPresentation`.  |
| `20-core-properties.pptx`   | Document metadata (Title / Author / Subject / Keywords).      |
| `21-showcase.pptx`          | The everything deck — 7 slides combining all of the above.    |

## How to inspect

1. Open each file in PowerPoint (Mac or Windows).
2. Watch for the "PowerPoint found a problem with content" dialog — that
   indicates the file is malformed.
3. For samples 16 and 17, run the slideshow (`F5`) to see transitions and
   animations.
4. For sample 20, check `File ▸ Info` for the metadata fields.
5. Open one or two samples in Keynote and Google Slides as a portability check.

The `samples/out/` directory is gitignored. Each PR-ready release should
regenerate the samples and inspect them before tagging.
