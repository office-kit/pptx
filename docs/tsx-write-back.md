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

Non-goals for this work:

- Lossless PPTX → TSX conversion of arbitrary decks. A `Presentation source={…}`
  deck keeps its original slides in the source `.pptx`; edits to them are out of
  scope until the DSL can address such shapes (see "Later").
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
  | { kind: 'unsupported'; shape?: ShapeRef; slide?: number; description: string };
```

Shape identity is the shape id: the editor loads `/editor/source`, the deck the
server built, so ids match; new shapes get fresh ids. Anything the diff cannot
express as one of the typed kinds (a new or deleted shape, a gradient, a run
format, an animation, slide order…) becomes `unsupported` with a human-readable
description — that is the hand-off to Claude, never a silent drop.

### 2. Shape → JSX element

`buildDeck` already compiles with the `jsxDev` transform, so
`getShapeJsxSources(shape)` returns the element chain for every DSL-made shape.
`BuildResult` gains `shapeSources: Record<number, Record<number, JsxSource[]>>`
(slide index → shape id → chain). The innermost entry is the element that created
the shape. A shape with no sources (source deck, `Raw`, `Fill` target) cannot be
written back.

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
- Number formatting: inches are written rounded to 0.001 in (≈ 914 EMU), angles to
  0.1°. Verification compares at that precision, and the editor reloads the
  rebuilt deck, so the snapped value is what the user then sees.

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

## Phases

1. **Plumbing and geometry**: `shapeSources` in `BuildResult`, `deck-diff` for
   bounds/rotation, `planPropEdit` for numbers, `source-sync` with verify and
   rollback, save endpoint returns written/pending. Browser test: drag and resize a
   `<Shape>` and a `<Text>`, check the TSX literal changed and no sidecar remains;
   drag a shape made in `rows.map(...)`, check it stays pending.
2. **Text, fill and stroke** for `Text` and `Shape`.
3. **Claude hand-off**: selection context in the prompt hook, pending list, the
   Apply with Claude Code notice, re-diff after Claude's turn.
4. **Structure**: insert a new shape as a JSX element in its slide, delete an
   element that makes exactly one shape, reorder slides written as literal
   `<Slide>` siblings.

Later (each needs DSL props first): run-level text formats, paragraph formats,
gradients, effects, transitions and animations, tables and charts, and a
`Target`-based edit element for slides that came from a source deck.

## Open questions

- Precision: is 0.001 in acceptable for geometry the user drags, or should the
  DSL accept EMU (`{ emu: 914400 }`) for exact positions?
- Should a refused write-back (computed value) be automatically routed to Claude,
  or only on the user's click? The draft says on click, so nothing runs without
  the user asking.
