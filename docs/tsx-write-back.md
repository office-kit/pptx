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

The three cases a literal edit cannot handle — a computed or shared value, an
element that makes several shapes, and a property without a DSL prop — each get a
deterministic TSX form in section 9, so **no editor change is left outside TSX**.
Claude is then optional: a reviewer that can turn mechanical overrides back into
intent ("make all bars 0.2 in taller" instead of three overrides), not the only
way an edit reaches the source.

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

With section 9 every edit has a TSX form, so the sidecar only holds edits
waiting on a failed build or a conflict. Claude's role is review and intent:

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

### 9. Nothing outside TSX

#### 9.1 Computed and shared values

The edit applies to the one element that made the shape; nothing else may move.

- `width={W}` where `W` is shared: only this element's attribute changes
  (`width={3.2}`). `W` and its other users are untouched.
- A numeric geometry expression keeps its relation and gains a delta (decided:
  delta is the default, not a literal):
  `x={col * 2.5}` → `x={col * 2.5 + 0.3}`. A later edit folds into the same
  trailing literal instead of stacking deltas. Colors, text and other non-numeric
  props are replaced by the literal.
- Text from data (`<Text>{row.label}</Text>`) is replaced on that element only if
  it makes one shape; otherwise 9.2 applies.

#### 9.2 One element, several shapes: instance keys and overrides

An element inside `rows.map(...)` or inside a component used several times makes
several shapes, so editing its attributes would change all of them. Each shape
instead gets a stable identity, and the edit becomes an override addressed to it.

- Identity is the **key path** along the shape's JSX chain. JSX already carries
  `key` (the runtime receives it and today ignores it); the DSL records it next to
  the source location. Write-back adds a key where a level is ambiguous:
  - a `.map` callback element without `key` gets `key={i}`, adding the index
    parameter to the callback if it has none;
  - a component call site used more than once on a slide gets a literal key from
    its first text prop or name (`<KpiCard key="revenue" …>`), deduplicated;
  - an element inside a component that makes more than one shape gets a static
    key (`<Text key="value" …>`).
- The override is an element in the slide, after the content it targets, using
  the `target` form from section 8 with a key path instead of an id:

  ```tsx
  {rows.map((row, i) => <Shape key={i} preset="rect" x={1} y={1 + i * 0.6} … />)}
  <Shape target={{ key: [1] }} y={1.75} />
  <Text target={{ key: ['revenue', 'value'] }} color="#C00000" />
  ```

- Key paths survive Claude's edits and data reordering as long as keys do; a
  target that matches no shape, or several, is a build error (never a silent
  retarget), so a stale override is reported instead of moving the wrong shape.
- Claude can later fold overrides into the data or the component, which is the
  "intent" step; until then the override is valid, typed TSX.

#### 9.3 Properties without a DSL prop: a coverage gate

Every mutating capability the editor exposes ends in a typed core setter
(`setShapeShadow`, `setParagraphSpacing`, `setSlideTransition`, …). The DSL gets a
typed prop for each, reusing the core option types so the type checker sees the
same discriminated unions the core does.

- A test (next to the editor's capability coverage test) maps every mutating
  editor capability to the DSL prop that expresses its result, and fails when one
  is missing. A new editor feature cannot ship without its TSX form, so the set of
  "editor can do it, TSX cannot say it" stays empty by construction.
- The deck diff's `unsupported` kind is then a bug signal, not a hand-off: the
  save fails loudly in development with the property name.
- The gate is complete before write-back is switched on (see Phases), so there
  is never a period in which the editor can make a change TSX cannot state, and
  no "disabled until supported" or "kept in the sidecar" interim is needed. Today
  that is about 120 mutating core APIs the editor calls (shape paint and effects,
  text and paragraph formats, tables, images, backgrounds, transitions,
  animations, comments, show settings, slide structure).
- The gate checks resulting state, not commands: the deck diff reads every
  property the editor can change through public getters, and each property must
  round-trip through a DSL prop (build → diff is empty).

#### 9.4 Parts of an imported deck the DSL does not model

Section 8 keeps unmodeled properties of an imported deck in the source package.
They are not edits, so they never need writing back, but they are invisible in
TSX. As DSL coverage grows (9.3), `import` writes more of them as props. Opaque
extension XML that has no typed model (vendor `extLst`, custom XML parts) stays
in the source package by design: it is preserved, not expressed.

## Phases

Write-back stays behind a development flag until phase 4 completes; the editor
keeps today's sidecar behaviour until then. When the flag is removed, every
change the editor can make already has a TSX form.

1. **Plumbing and geometry**: `shapeSources` and key paths in `BuildResult`,
   `deck-diff` for bounds/rotation, `planPropEdit` (literal replace, attribute
   insert, numeric delta), `source-sync` with verify and rollback.
2. **Structure** (section 7): added, pasted, duplicated and deleted shapes; new,
   deleted and reordered slides.
3. **Instances** (9.2): `key` recording, key insertion, `target={{ key }}`
   overrides.
4. **Full coverage** (9.3): DSL props for every property the editor changes,
   category by category, with the round-trip gate; then remove the flag.
5. **`target` form and `office-pptx import`** (section 8), with round-trip tests
   on the sample decks in `samples/`.
6. **Claude as reviewer**: selection context in the prompt hook and a "tidy
   overrides" action that asks Claude to fold overrides into data or components.

Layouts and masters follow the same pattern with a `<Layout target>` element,
which the Design tab's layout commands need before they can be written back.
