# TSX authoring reference

Use the installed `@office-kit/pptx-dsl` declarations for exact prop types. This
runtime is independent of React and Vue; there is no DOM or CSS layout.
The initialized project already configures `jsxImportSource` and type checking.

## Native slides

Default-export a `Presentation` from `deck.tsx`. Coordinates and dimensions are
inches, font sizes and stroke widths are points. Default size is 16:9,
approximately 13.333 × 7.5 inches. Child order is drawing order.

```tsx
import { Presentation, Slide, Text, Shape, Chart, Table } from '@office-kit/pptx-dsl';

const accent = '#E5481F';

export default (
  <Presentation>
    <Slide background="#15171C" notes="Introduce the quarterly results.">
      <Shape preset="rect" x={0.9} y={2.5} width={0.15} height={1.6} fill={accent} stroke={false} />
      <Text x={1.3} y={2.5} width={10.5} height={1.2} size={44} bold color="#FFFFFF">
        Quarterly business review
      </Text>
      <Text x={1.3} y={4} width={10} height={0.6} size={20} color="#B4B9C4">
        Illustrative figures — replace with actual results
      </Text>
    </Slide>
    <Slide background="#FFFFFF">
      <Text x={0.9} y={0.5} width={11.5} height={0.8} size={30} bold>
        Revenue by quarter
      </Text>
      <Chart
        x={0.9}
        y={1.7}
        width={11.5}
        height={5}
        spec={{
          kind: 'column',
          categories: ['Q1', 'Q2', 'Q3', 'Q4'],
          series: [{ name: 'Revenue', values: [120, 180, 240, 300], color: accent }],
        }}
      />
    </Slide>
    <Slide background="#FFFFFF">
      <Text x={0.9} y={0.5} width={11.5} height={0.8} size={30} bold>
        Next actions
      </Text>
      <Table
        x={0.9}
        y={1.7}
        width={11.5}
        height={2.4}
        rows={[
          ['Owner', 'Action', 'Due'],
          ['Aiko', 'Review the hosting contract', 'Nov 15'],
          ['Ben', 'Ship annual billing', 'Dec 1'],
        ]}
        columnWidths={[2.2, 7, 2.3]}
        cellStyle={{ format: { size: 16, color: '#15171C' } }}
        headerStyle={{ fill: '#15171C', format: { color: '#FFFFFF', bold: true } }}
        stripeFill="#F3F4F7"
      />
    </Slide>
  </Presentation>
);
```

Use functions, fragments, arrays, conditions and `.map()` to compose content.
Bounds (`x`, `y`, `width`, `height`) are required for new visual objects.

| Element        | Inputs                                                                                                                            |
| -------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| `Presentation` | Optional `source` bytes, new-deck `size`, `theme`, `mode`.                                                                        |
| `Slide`        | `background`, `notes`, `layout`, source `from` or edit `target`.                                                                  |
| `Text`         | Literal children or rich core `ParagraphSpec[]` in `paragraphs`; core `TextFormat` props such as `size`, `font`, `bold`, `color`. |
| `Shape`        | `preset`, solid/gradient `fill`, `stroke`, rotation, effects, optional text.                                                      |
| `Image`        | Byte `data`, optional `format` and `fit`. Load local bytes with `readFile(new URL('./image.png', import.meta.url))`.              |
| `Chart`        | Core `ChartSpec` in `spec`, including categories and series.                                                                      |
| `Table`        | String `rows`, `columnWidths`, `rowHeights`, `cellStyle`, `headerStyle`, `stripeFill`.                                            |
| `Fill`         | Existing shape `target`, replacement text children, optional `format`.                                                            |
| `Remove`       | Existing shape `target`.                                                                                                          |
| `Raw`          | Public core API callback with optional explicit scope.                                                                            |

## Existing PPTX files

Run `npx --no-install office-pptx inspect template.pptx` in the project first.
Use its actual slide and shape references, not guesses such as "Title 1".

```tsx
import { readFile } from 'node:fs/promises';
import { Presentation, Slide, Fill } from '@office-kit/pptx-dsl';

const source = await readFile(new URL('./template.pptx', import.meta.url));

export default (
  <Presentation source={source} mode="edit">
    <Slide target={{ index: 0 }}>
      <Fill target={{ name: 'Title 1' }}>Quarterly business review</Fill>
    </Slide>
  </Presentation>
);
```

Change `Title 1` to a real target from inspection. Slide references accept a
zero-based `index`, exact `title` or `part`; layouts accept exact `name`, `type`
or `part`; shapes accept exact `name`, `id` or `placeholder` with `type`/`idx`.
Missing and ambiguous targets fail.

`mode="edit"` is the default with source bytes. It keeps unmentioned slides and
existing masters, layouts and unknown package parts. `Fill` intentionally replaces
the selected shape's text. A rebuild loads the source fresh; write to a separate
output file so edits are not reapplied to their own output.

Use `mode="compose"` with `<Slide from={{ index: 0 }}>` to duplicate source slides
into a new sequence, only when that is the intent. Compose removes the original
slide sequence after compilation. References always refer to the original source
sequence. Source decks retain their dimensions; `size` applies only to new decks.
Do not use cross-presentation imports as a lossless import path.

## Raw escape hatch

Use this for a supported core capability with no typed DSL prop:

```tsx
import { setSlideTransition } from '@office-kit/pptx';
import { Presentation, Slide, Text, Raw } from '@office-kit/pptx-dsl';

export default (
  <Presentation>
    <Slide>
      <Text x={1} y={1} width={10} height={1} size={32}>
        Welcome
      </Text>
      <Raw
        scope="slide"
        apply={({ slide }) => {
          setSlideTransition(slide, { effect: 'fade', speed: 'med' });
        }}
      />
    </Slide>
  </Presentation>
);
```

Scopes `presentation`, `slide`, and `shape` require the corresponding enclosing
context. Without a scope, contextual handles are optional. Callbacks can be async;
they run in declaration order after ordinary construction and before compose mode
removes originals. Core geometry uses `inches(...)`/`pt(...)` helpers, unlike the
DSL's numeric inch props. See [core API details](core-api.md) as needed.

Full PPTX expressiveness is a goal, not current typed coverage. The static preview
does not play transitions or media. Preserve unsupported existing content even
when the preview cannot render it; do not recreate a template from visible objects.

## Organize for focused revisions

Keep `deck.tsx` as a small composition of imported slide functions:

```tsx
import { Presentation } from '@office-kit/pptx-dsl';
import { Cover } from './slides/cover.tsx';
import { Revenue } from './slides/revenue.tsx';

export default (
  <Presentation>
    <Cover />
    <Revenue />
  </Presentation>
);
```

Each function returns a `Slide`. Use `theme.ts` for shared palette/typography
values; import only what the slide needs. Use descriptive filenames that stay
stable when slides move. Patch text, props or the relevant data item for local
changes; change a shared value only when all its consumers should change.
Existing inline slides and data-driven compositions remain valid.
