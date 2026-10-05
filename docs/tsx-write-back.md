# TSX write-back: one source for the editor and Claude Code

Status: design, not implemented. Branch `feat/tsx-write-back`, stacked on
`feat/pptx-editor` (PR #287).

## Problem

The `office-pptx dev` preview has two ways to change a deck, and they write to
different places:

| Who                           | Writes to                                                        |
| ----------------------------- | ---------------------------------------------------------------- |
| Claude Code (terminal pane)   | The TSX source (`deck.tsx` and the files it imports)             |
| The visual editor (`/editor`) | `.office-kit/<entry>.editor.zip`, a whole edited `.pptx` sidecar |

The sidecar records the source hash it was based on. Once Claude changes the
source, the hashes differ and the user must pick "keep my edits" or "use the
source" — one side's work is thrown away. Claude never sees what was edited
visually, and the visual edits never become typed, reviewable code.

## Goal

The TSX source is the only source of truth. A visual edit becomes a source edit
whenever the change can be stated in the DSL's typed props; anything else goes
to Claude Code with the selection and the intended change, and Claude edits the
source. The sidecar shrinks to "edits not yet in the source" and is empty in the
common case.

Invariant: **every slide and every shape the editor can touch is a JSX element
in the source.** There is no "this object only exists in the PPTX" state. Two
paths used to create one, and both are closed:

1. Objects the editor creates (insert, paste, duplicate, new slide) are written
   back as new JSX elements, not left in the sidecar (section 7).
2. A deck that starts from an existing `.pptx` is imported into TSX that names
   every slide and shape as an element (section 8). Properties the DSL cannot
   express yet stay in the source package by reference, so nothing is lost, but
   the object itself is always addressable in TSX.

What can still go to Claude is a change that is in TSX but not mechanically
writable: a computed or shared value, an element that makes several shapes
(`rows.map`, a reused component), or a property the DSL has no prop for yet.

Non-goals for this work:

- Writing into shared components or computed values. Those go to Claude.

## Design

### Overview

```
editor (in-memory deck) ──save──▶ server
                                   │ 1. diff: edited deck vs deck built from source
                                   │ 2. plan: each change → a TSX prop edit at the
                                   │    shape's JSX element (TypeScript AST)
                                   │ 3. write all planned edits, rebuild, verify
                                   │    the rebuilt deck matches the edited deck
                                   │    on the changed properties; else roll back
                                   ▼
                     ┌─ all changes written ──▶ sidecar cleared, editor reloads
                     └─ some remain ──────────▶ sidecar keeps the remainder,
                                                 "Apply with Claude Code" offered
```

The editor keeps editing an in-memory deck exactly as today, so typing, dragging
and undo stay instant. Write-back runs at the existing autosave point.

### 1. Deck diff (`packages/dev/src/deck-diff.ts`)

Compares two `PresentationData` through public getters and returns typed changes:

```ts
type ShapeRef = { slide: number; shapeId: number };
type DeckChange =
  | {
      kind: 'bounds';
      shape: ShapeRef;
      value: { x: number; y: number; width: number; height: number };
    } // inches
  | { kind: 'rotation'; shape: ShapeRef; value: number }
  | { kind: 'text'; shape: ShapeRef; value: string } // plain text of a single-run body
  | { kind: 'fill'; shape: ShapeRef; value: api.Color | null }
  | { kind: 'stroke'; shape: ShapeRef; value: { color: api.Color; width: number } | false }
  | { kind: 'added'; shape: ShapeRef } // created, pasted or duplicated in the editor
  | { kind: 'removed'; shape: ShapeRef }
  | { kind: 'slides'; order: (number | 'new')[] } // source slide indices, in the edited order
  | { kind: 'unsupported'; shape?: ShapeRef; slide?: number; description: string };
```

Shape identity is the shape id: the editor loads `/editor/source`, the deck the
server built, so ids match; new shapes get fresh ids. Anything the diff cannot
express as one of the typed kinds (a gradient, a run format, an animation…)
becomes `unsupported` with a human-readable
description — that is the hand-off to Claude, never a silent drop.

### 2. Shape → JSX element

`buildDeck` already compiles with the `jsxDev` transform, so
`getShapeJsxSources(shape)` returns the element chain for every DSL-made shape.
`BuildResult` gains `shapeSources: Record<number, Record<number, JsxSource[]>>`
(slide index → shape id → chain). The innermost entry is the element that created
the shape. Shapes a source deck brought along get their element from the
`target` form in section 8, so they have sources too. Only `Raw`-made shapes have
none; `Raw` is the escape hatch and its output is Claude's to edit.

### 3. Prop edit planner (`@office-kit/pptx-dsl/source-edit`)

A pure function next to `planTextEdit`, with the same refusal style:

```ts
planPropEdit({
  file: SourceFile,
  anchor: TextEditAnchor,           // the element's `<`
  prop: 'x' | 'y' | 'width' | 'height' | 'rotation' | 'fill' | 'stroke' | 'text',
  value: number | string | false | { color: string; width: number },
}): { ok: true; change } | { ok: false; reason: 'not-element' | 'computed' | 'spread' | 'unsupported-element' }
```

Rules:

- The anchor must be a JSX element of a known DSL kind (`Text`, `Shape`, `Image`,
  `Line`, `Table`, `Chart`, `Media`, `Group`), else `unsupported-element` (e.g. a
  component call site — writing into the component is Claude's job).
- An existing attribute is replaced only if its initializer is a literal (number,
  negative number, string, `false`, or an object literal of literals). An
  identifier, call, conditional or template with substitutions is `computed`.
- A missing attribute is inserted, unless the element has a spread attribute
  (`{...props}`) that could also set it (`spread`).
- `text` follows `planTextEdit`'s existing rules, anchored to the element.
- Number formatting follows the OOXML wire values exactly. Geometry is
  `ST_Coordinate` (whole EMU) and rotation `ST_Angle` (whole 60000ths of a
  degree); the DSL takes inches and degrees and rounds them with `inches()` /
  the rotation setter. The planner writes the shortest decimal that rounds back
  to the exact wire value (914400 EMU → `1`, 1371600 → `1.5`; six decimals always
  suffice for EMU, five for angles), and verification compares wire values for
  equality, so a write-back never moves a shape by even one EMU.

The TypeScript AST is already used by `planTextEdit`; `typescript` stays an
optional peer of the DSL, so the runtime bundle does not grow.

### 4. Apply, verify, roll back (`packages/dev/src/source-sync.ts`)

Generalizes `text-edit.ts`:

1. Group planned edits by file; refuse if any file changed since the build the
   diff used.
2. Write all files, rebuild once.
3. Diff the rebuilt deck against the edited deck. Success means: every planned
   change now matches, and nothing outside the planned shapes changed (a literal
   reused by two shapes, e.g. a shared `const w = 3`, is caught here because we
   only write literals in place — but a component instantiated twice is not, so
   this check is what guards it).
4. On failure restore every file and rebuild; the change stays in the sidecar.

The save request returns `{ written: DeckChange[], pending: DeckChange[] }`.

### 5. Hand-off to Claude Code

- The editor already posts `editor-focus` (slide index) to the preview page. It
  grows to carry selected shape ids, their JSX sources and the text selection.
  The `UserPromptSubmit` hook (`/terminal/context`) already injects focus context
  into Claude's prompt; it adds the selection and, when the sidecar is non-empty,
  the `pending` list ("shape `Revenue chart` (deck.tsx:42): fill → gradient …").
- The editor shows a notice for pending changes with **Apply with Claude Code**,
  which types a prepared instruction into the terminal (the user still presses
  Enter). The existing `Stop` hook already rebuilds and type-checks; after Claude
  finishes, the server re-diffs and clears whatever the source now matches.
- If the source changes while pending edits exist, the server re-diffs the
  sidecar against the new build. Only when a pending property was also changed by
  the source differently is it a conflict, shown per change rather than as
  "keep everything / discard everything".

### 6. Undo

The editor's undo stack holds deck snapshots. Undoing after a write-back yields a
deck that differs from the source, which the next save writes back like any other
edit, so undo needs no special case. The project history (`/history`) already
snapshots the source files per turn, so source write-backs appear there too.

### 7. Structural write-back

- **Added shape**: a new element is inserted as the last child of the slide's
  `<Slide>` element (z-order follows document order), with the props the diff
  can state: `<Shape preset="rect" x={…} … />`, `<Text …>…</Text>`,
  `<Image data={…} />`. Image bytes are written next to the deck
  (`assets/<hash>.png`) and imported, never inlined. If the slide is not a
  literal `<Slide>` element (e.g. made by `slides.map`), the addition is
  pending for Claude.
- **Removed shape**: an element that makes exactly one shape is deleted; a
  shape from a source deck becomes `<Remove target={{ id }} />` (section 8).
- **Slide order, new and deleted slides**: literal `<Slide>` siblings are
  reordered, inserted (`<Slide layout={{ name }}>`) or deleted.
- Paste and duplicate are additions with the copied props.

### 8. Importing an existing `.pptx` as TSX

`office-pptx import deck.pptx` writes `deck.tsx`:

```tsx
const source = await readFile(new URL('./deck.pptx', import.meta.url));
export default (
  <Presentation source={source} mode="compose">
    <Slide from={{ index: 0 }}>
      <Text target={{ id: 2 }} x={0.5} y={0.4} width={9} height={1.2}>
        Q3 business review
      </Text>
      <Shape target={{ id: 5 }} x={0.5} y={2} width={4} height={3} fill="#E8EEF7" />
      <Image target={{ id: 7 }} x={5} y={2} width={4} height={3} />
    </Slide>
    …
  </Presentation>
);
```

- Every slide is a `<Slide from>` in compose mode, so slides can be reordered,
  added and deleted as elements.
- Every shape is an element with a new `target` form, the same convention
  `<Slide target>` already uses: with `target`, `Shape` / `Text` / `Image` /
  `Table` / `Chart` / `Line` / `Group` / `Media` edit the existing shape instead of
  creating one. The props the DSL can express are written out explicitly, so the
  TSX shows the deck's real geometry, text and colors and type-checks them; every
  other property (effects, run formats the DSL lacks, unknown parts) stays in the
  source package untouched. This keeps the CLAUDE.md rule that unknown parts are
  never discarded or flattened.
- A shape with no element in a compose-mode `<Slide from>` is removed from the
  output only through `<Remove>`; an element-less shape is a build warning, so the
  invariant holds even after hand edits.
- As DSL coverage grows, `import` writes more props and the source package
  matters less; a deck whose every property is expressible can drop `source`.

Open points for this section: `Fill target` overlaps `Text target` (one way to
do one thing — `Fill` would be deprecated in favour of `Text target`), and
`<Slide from>` must keep shape ids stable, which needs a test.

## Phases

1. **Plumbing and geometry**: `shapeSources` in `BuildResult`, `deck-diff` for
   bounds/rotation, `planPropEdit` for numbers, `source-sync` with verify and
   rollback, save endpoint returns written/pending. Browser test: drag and resize a
   `<Shape>` and a `<Text>`, check the TSX literal changed and no sidecar remains;
   drag a shape made in `rows.map(...)`, check it stays pending.
2. **Structure** (section 7): added, pasted, duplicated and deleted shapes; new,
   deleted and reordered slides.
3. **Text, fill and stroke** for `Text` and `Shape`.
4. **`target` form and `office-pptx import`** (section 8), with round-trip tests
   on the sample decks in `samples/`.
5. **Claude hand-off**: selection context in the prompt hook, pending list, the
   Apply with Claude Code notice, re-diff after Claude's turn.

Later (each needs DSL props first): run-level text formats, paragraph formats,
gradients, effects, transitions and animations, tables and charts, layouts and
masters (the Design tab's layout commands need a `<Layout target>` element before
they can be written back).

## Open questions

- Should a refused write-back (computed value) be automatically routed to Claude,
  or only on the user's click? The draft says on click, so nothing runs without
  the user asking.
