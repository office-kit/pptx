# @office-kit/pptx-dsl

Start with [Create slides with AI](https://office-kit.github.io/pptx/docs/authoring)
for installation and preview, then see the
[TSX reference](https://office-kit.github.io/pptx/docs/tsx).
Install the Claude Code skill to let the agent handle project setup, or run
`npx --yes @office-kit/pptx-dev@latest init my-slides` to create a project yourself.
For an existing TypeScript project, install `@office-kit/pptx-dsl` and
`@office-kit/pptx` from npm.

Typed, declarative PPTX authoring in TSX. This package uses its own JSX
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
| `Table`        | Rows of strings, rich cells or merges; sizes; cell styles and `styleCell` |
| `Fill`         | Existing shape target, replacement text, optional format and `autoFit`    |
| `Remove`       | Existing shape target                                                     |
| `Raw`          | Deferred callback receiving the presentation and enclosing slide/shape    |

Bounds are required for newly created visual objects. Shape styles support solid
and gradient fills, stroke, rotation, shadow, glow and click actions. `Text`
accepts the core `TextFormat` properties directly, plus alignment, anchor,
text auto-fit, `bullets` and `paragraphSpacing` (both apply to every paragraph; a
newline in the text starts one). A `Shape` with `text` takes `align` and `anchor`.

Each entry of `paragraphs` is a core `ParagraphSpec` that also takes its own `bullet`
and `level`. The paragraph's `bullet` wins over `bullets`, so `bullet: 'none'` keeps a
heading line out of the list. `level` is the `TextLevel` type, `0` to `8`, and nests
the list; `tsc` rejects any other number.

```tsx
<Text
  x={1}
  y={1}
  width={8}
  height={3}
  bullets="bullet"
  paragraphs={[
    { runs: [{ text: 'What changed' }], bullet: 'none' },
    { runs: [{ text: 'Churn fell to 4.2%' }] },
    { runs: [{ text: 'Driver: simpler setup' }], level: 1, bullet: { char: '–' } },
  ]}
/>
```

A `Line` runs from (`x1`, `y1`) to (`x2`, `y2`) instead of taking bounds. A
`Group` holds two or more visual elements and may contain other groups; the
result moves and resizes as one object in presentation apps.

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

A cell in `rows` is a string or a cell object, `{ text }` or `{ paragraphs }`, never
both or neither. `paragraphs` takes core `ParagraphSpec[]` for more than one run or
paragraph in a cell. The merged cell style is the base of every run, so a run states
only what differs, and the style's `align` applies unless a paragraph has its own.
`styleCell` receives a rich cell's `value` as its run texts joined, one line per
paragraph.

`colSpan` and `rowSpan` on a cell object make it the top-left of a merge. `rows` stays
a full rectangular grid, so a column index means the same in every row: each position
the merge covers is written as `''`, and anything else there is an error because
a covered cell's text is never shown. The merged block takes its fill and borders from
its top-left cell (this is how the preview renders it), so covered positions get no
style and `styleCell` is not called for them. Spans that leave the grid or overlap throw the core error.

```tsx
<Table
  x={1}
  y={1}
  width={8}
  height={3}
  rows={[
    [{ text: 'Plan', rowSpan: 2 }, { text: 'Effect', colSpan: 2 }, ''],
    ['', 'Count', 'Change'],
    [
      'FAQ',
      '372',
      { paragraphs: [{ runs: [{ text: '+26% ' }, { text: 'QoQ', format: { bold: true } }] }] },
    ],
  ]}
  cellStyle={{ format: { size: 14 } }}
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
`Fill` takes the same `autoFit` values as `Text` (`"normal"` shrinks text that is
longer than the placeholder); omitted, the shape keeps the template's setting.
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

## Embedding `compile()` in your own build

An `office-pptx init` project comes with a build that records source locations and
a `check` script that runs `tsc --noEmit`. A host that calls `compile()` from its own
build (a server compiling generated TSX, a bundler plugin) has to arrange both.

**Build with the dev JSX transform to get source locations.** `compile()` prefixes
an error with `file:line:col <Element>:` for each enclosing element, but only when the
module was built with the dev runtime, which records where each element was written.
`office-pptx` always builds that way.

| Tool       | Setting                                        |
| ---------- | ---------------------------------------------- |
| TypeScript | `"jsx": "react-jsxdev"` instead of `react-jsx` |
| esbuild    | `jsx: 'automatic', jsxDev: true`               |

```text
react-jsx      Shape reference matched 0 shapes: {"name":"Titel 1"}
react-jsxdev   deck.tsx:6:7 <Slide>: deck.tsx:7:9 <Fill>: Shape reference matched 0 shapes: {"name":"Titel 1"}
```

**Type-check authored modules with `tsc` in strict mode.** The runtime does not
validate prop names, so `<Table cellFills={...}>` compiles and the unknown prop is
ignored. Type checking is what rejects it, along with a misspelt literal such as
`anchor="ctr"`. A transpile-only build (esbuild, swc) skips that check, and a prop
passed through an object spread escapes it even in `tsc`.

**Compile and save.** `compile(root)` takes the root `Presentation` element (a
module's default export) and resolves to the core presentation; `savePresentation`
from `@office-kit/pptx` turns that into `.pptx` bytes. Template bytes go in through
`<Presentation source={bytes}>`, from wherever the host holds them.

```ts
import { savePresentation } from '@office-kit/pptx';
import { compile } from '@office-kit/pptx-dsl';

const { default: root } = await import(builtModuleUrl);
const bytes = await savePresentation(await compile(root));
```

## From a shape back to its source

`getShapeJsxSources(shape)` returns the source locations of the elements that
were being evaluated when `compile()` created the shape, outermost first. It needs
the dev JSX transform described above. A component call contributes its call site,
so a shape a prebuilt component made (whose own elements carry no location) still
leads back to `<Headline … />` in the slide file, and every shape a component returns
through a fragment gets that location. A `Group` has its own location; its members
keep theirs.

```ts
import { getShapeId, getSlideShapes, getSlides } from '@office-kit/pptx';
import { compile, getShapeJsxSources } from '@office-kit/pptx-dsl';

const presentation = await compile(root);
const sources = getSlides(presentation).map((slide) =>
  getSlideShapes(slide).map((shape) => ({
    id: getShapeId(shape),
    sources: getShapeJsxSources(shape),
  })),
);
```

- `lineNumber` and `columnNumber` are 1-based and point at the element's `<`, as
  TypeScript's `react-jsxdev` and esbuild's `jsxDev` emit them. `fileName` is the
  name the transform was given.
- It is `undefined` for shapes built with the production runtime, shapes a source
  deck or `Slide from` brought along, targets of `Fill`, and shapes `Raw` made.
- It answers for the presentation `compile()` returned, until that presentation's
  slides or shapes are added, removed or regrouped. Nothing is written into the PPTX.
  Shape ids survive `savePresentation` and `loadPresentation`, so a map from
  slide index and `getShapeId` to sources, built before saving, stays valid for the
  saved file.

### Editing a literal in place

`@office-kit/pptx-dsl/source-edit` holds the text-edit rules the dev preview's
double-click edit uses, as pure functions over source strings. It needs
`typescript`, an optional peer dependency.

`planTextEdit` finds the one source literal that renders `before` and returns the
file with it replaced by `after`: JSX text (whitespace collapsed the way JSX renders
it), an attribute string, a string literal or a template literal without
substitutions. JSX gets `{"…"}`, so any text is safe. It refuses with a reason code:

| Reason                | Meaning                                                                              |
| --------------------- | ------------------------------------------------------------------------------------ |
| `invalid`             | `before` is blank, `before` equals `after`, or either exceeds `TEXT_EDIT_MAX_LENGTH` |
| `not-unique-on-slide` | `before` does not occur exactly once in `slideText`                                  |
| `not-found`           | No literal matches: the text is computed, formatted, or in another file              |
| `ambiguous`           | More than one literal matches                                                        |

Pass an `anchor` (a `getShapeJsxSources` entry, with `fileName` as `path`) to search
only the element that starts there; two literals inside that element are still
`ambiguous`. Only `files` are searched, so leave shared components out to keep their
literals untouched.

After rebuilding with the change, `verifyTextEdit` confirms that the only difference
is `before` → `after` on the edited slide, or returns `slide-count-changed`,
`other-slides-changed` or `text-mismatch`. Writing, rebuilding and rolling back are
the caller's.

```ts
import { planTextEdit, verifyTextEdit } from '@office-kit/pptx-dsl/source-edit';

const plan = planTextEdit({ files, slideText, before, after, anchor });
if (plan.ok) {
  const next = await rebuild(plan.change); // { slides: svg[], slideTexts: string[] }
  const verdict = verifyTextEdit({ previous, next, slide, before, after });
}
```

## Current coverage

This is an initial DSL, not complete OOXML coverage. Typed master/layout creation,
animations, transitions, connectors, groups and media playback still require additional declarative elements or core work. Raw can invoke
available core capabilities but does not count as typed declarative support.
Cross-presentation imports are not exposed because the current core import API
can discard unsupported relationships. Do not use it as a lossless import path.
