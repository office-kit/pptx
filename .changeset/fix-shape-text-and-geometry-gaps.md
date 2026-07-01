---
'pptx-kit-preview': patch
---

Fix several text-on-shape and shape-geometry rendering gaps found by comparing
output against LibreOffice ground truth on a corpus of realistic, multi-feature
decks:

- **Preset pattern fills** (`pct5`–`pct90`, `smGrid`/`lgGrid`, and the
  horizontal/vertical/diagonal hatch families) now match LibreOffice's actual
  substitution for these fills — a density-scaled diagonal hatch for the
  percentage family, and correctly differentiated tile pitches for the "small"
  vs "large" grid variants — instead of a uniformly-dense ordered-dither screen.
- **Multi-column text bodies** (`numCol` with `noAutofit`/`spAutoFit`) now wrap
  into a new row of columns once the last column also overflows, instead of
  piling all remaining text into the final column forever.
- **`u="wavy"`** (and `wavyDbl`/`wavyHeavy`) now renders as an actual wavy
  underline in the SVG/raster path (drawn as an explicit path — resvg has no
  `text-decoration-style` support) and as real CSS in the browser path.
- **Rotated + vertically-flipped shape text** no longer renders upside-down;
  PowerPoint adds a compensating 180° turn to the text specifically for
  `flip.vertical`, independent of `flip.horizontal`.
- **Pie/doughnut data labels** now join as `<category> — <value/percent>`
  (e.g. "Web — 48%"), matching PowerPoint/LibreOffice's order instead of the
  reverse.
- **Overlapping (non-stacked) area charts** now paint series back-to-front so
  the first-authored series stays on top, matching PowerPoint; the
  category-axis title no longer collides with the tick-label row.
- **Table row/column banding** now uses a pale tint (not a near-solid accent
  color), alternates between two tints across every body row (previously every
  other row was left unshaded), and starts the alternation at the first body
  row rather than one row late.
