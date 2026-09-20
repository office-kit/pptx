# @office-kit/pptx-dsl

Start with [Create slides with AI](https://office-kit.github.io/pptx/docs/authoring)
for installation and preview, then see the
[TSX reference](https://office-kit.github.io/pptx/docs/tsx).
Install the Claude Code skill to let the agent handle project setup, or run
`npx --yes @office-kit/pptx-dev@latest init my-slides` to create a project yourself.
For an existing TypeScript project, install `@office-kit/pptx-dsl` and
`@office-kit/pptx` from npm.

Typed, declarative PowerPoint authoring in TSX. This package uses its own JSX
runtime; React and Vue are not dependencies. It creates native editable PPTX
objects through `@office-kit/pptx`. Rendering belongs to `@office-kit/pptx-preview`.

```tsx
import { Presentation, Slide, Text, Chart } from '@office-kit/pptx-dsl';

export default (
  <Presentation>
    <Slide background="#FFFFFF">
      <Text x={0.9} y={0.5} width={11.5} height={0.8} size={30} bold>
        Quarterly revenue
      </Text>
      <Chart
        x={0.9}
        y={1.7}
        width={11.5}
        height={5}
        spec={{
          kind: 'column',
          categories: ['Q1', 'Q2'],
          series: [{ name: 'Revenue', values: [120, 180] }],
        }}
      />
    </Slide>
  </Presentation>
);
```

Configure TypeScript with `jsx: "react-jsx"`,
`jsxImportSource: "@office-kit/pptx-dsl"`, and `moduleResolution: "Bundler"`.
The `@office-kit/pptx-dev` CLI supplies project initialization, export and watch
preview. Alternatively, call `compile(root)` and `savePresentation(result)`.
The DSL itself is browser-safe and performs no filesystem access.

Coordinates and dimensions (`x`, `y`, `width`, `height`) are inches. Font size
and stroke width are points. Child order is drawing order. Functions, fragments,
arrays, conditions and `.map()` compose elements using ordinary TypeScript.
Layout is explicit: CSS and automatic UI layout are not implemented.

## Elements

| Element        | Input                                                                     |
| -------------- | ------------------------------------------------------------------------- |
| `Presentation` | Optional source PPTX bytes or `size`, theme, `mode`                       |
| `Slide`        | Source `from` or edit `target`, layout, background, notes                 |
| `Text`         | Literal children or core `ParagraphSpec[]` via `paragraphs`, text format  |
| `Shape`        | Core preset geometry, fill, stroke, effects, optional text                |
| `Line`         | End points `x1` `y1` `x2` `y2`, `color`, `width`                          |
| `Group`        | Groups the shapes its children create; optional `name`                    |
| `Image`        | Bytes in `data`, optional format and fit                                  |
| `Media`        | `kind` `video` / `audio` with bytes in `data`, or `online` with a `url`   |
| `Chart`        | Complete core `ChartSpec` via `spec`                                      |
| `Table`        | String rows, column and row sizes, cell/header/stripe styles, `styleCell` |
| `Fill`         | Existing shape target, replacement text, optional format                  |
| `Remove`       | Existing shape target                                                     |
| `Raw`          | Deferred callback receiving the presentation and enclosing slide/shape    |

Bounds are required for newly created visual objects. Shape styles support solid
and gradient fills, stroke, rotation, shadow, glow and click actions. `Text`
accepts the core `TextFormat` properties directly, plus alignment, anchor,
text auto-fit, `bullets` and `paragraphSpacing` (both apply to every paragraph; a
newline in the text starts one). A `Shape` with `text` takes `align` and `anchor`.

A `Line` runs from (`x1`, `y1`) to (`x2`, `y2`) instead of taking bounds. A
`Group` holds two or more visual elements and may contain other groups; the
result moves and resizes as one object in PowerPoint.

A table cell style sets `fill`, `format`, `anchor`, `align` and `borders` (per
side, `width` in points). Styles merge in this order, later ones winning per
field: `cellStyle`, `headerStyle` (row 0), `stripeFill` (even body rows), then
`styleCell`, a callback that receives `{ row, column, value }`:

```tsx
<Table
  x={1}
  y={1}
  width={8}
  height={3}
  rows={rows}
  cellStyle={{ borders: { bottom: { color: '#D5D9E0', width: 0.75 } } }}
  styleCell={({ row, value }) =>
    row > 0 && value === 'At risk' ? { fill: '#D64545', format: { color: '#FFFFFF' } } : undefined
  }
/>
```

`Media` embeds a video or audio clip from bytes, or links an online video by
URL (YouTube page URLs are rewritten to the embed form). `poster` supplies the
image shown before playback; a neutral play-button poster is used otherwise.
It takes bounds and `name` but no shape styles: the clip's picture already owns
its click action.

```tsx
<Media kind="video" data={clipBytes} poster={posterBytes} x={1} y={1.5} width={8} height={4.5} />
<Media kind="online" url="https://youtu.be/dQw4w9WgXcQ" x={1} y={1.5} width={8} height={4.5} />
```

## Existing presentations

```tsx
import { readFile } from 'node:fs/promises';
import { Presentation, Slide, Fill } from '@office-kit/pptx-dsl';

const source = await readFile(new URL('./template.pptx', import.meta.url));
export default (
  <Presentation source={source} mode="edit">
    <Slide target={{ index: 0 }}>
      <Fill target={{ name: 'Title 1' }}>Q3 business review</Fill>
    </Slide>
  </Presentation>
);
```

`mode="edit"` keeps the existing slide sequence and edits only explicit targets.
It is the default when a source is supplied. `mode="compose"` replaces the slide sequence with the declared
slides; use `<Slide from={{index: 0}}>` to duplicate an original slide. References
always refer to the source sequence, and each compilation loads the source fresh.
A source deck is not reconstructed from DSL-understood fields: the core package
retains existing parts. Editing text intentionally replaces that shape's text.
`size` applies only to new decks; source decks retain their original dimensions.
`theme` accepts the core theme overrides and follows its first-theme semantics.

Slide references accept a zero-based index, exact title, or part name. Layout
references accept exact name, type, or part name. Shape targets accept exact name,
id, or `{placeholder: {type, idx}}`. Missing and ambiguous matches are errors.
`office-pptx inspect template.pptx` lists usable references. Layouts and masters
from the template remain in the package; selecting a layout retains inheritance.

## Escape hatch

```tsx
import { setShapeRotation } from '@office-kit/pptx';
import { Shape, Raw } from '@office-kit/pptx-dsl';

const element = (
  <Shape preset="rect" x={1} y={1} width={2} height={1}>
    <Raw
      scope="shape"
      apply={({ shape }) => {
        setShapeRotation(shape, 15);
      }}
    />
  </Shape>
);
```

Raw callbacks run in declaration order after normal construction and before the
compose mode removes original slides. They can be async. They use public core
APIs. Set `scope="shape"`, `scope="slide"` or `scope="presentation"` for required,
context-specific handles and a nesting check. Without scope, contextual handles
are optional. Text also
accepts Raw children. An enclosing shape does not accept new Slide-level objects.

## Current coverage

This is an initial DSL, not complete OOXML coverage. Typed master/layout creation,
animations, transitions, connectors, groups, media playback, and structured table
cells still require additional declarative elements or core work. Raw can invoke
available core capabilities but does not count as typed declarative support.
Cross-presentation imports are not exposed because the current core import API
can discard unsupported relationships. Do not use it as a lossless import path.
