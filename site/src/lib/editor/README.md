# @office-kit/pptx — Editor

A PowerPoint-style editing UI built **entirely on the `@office-kit/pptx` public
API**, in Svelte 5. SvelteKit serves it at `/editor`; `@office-kit/pptx-dev`
bundles the same components into its local development preview. The development
host persists edits and resolves source conflicts; see
[the development-tool README](../../../../../packages/dev/README.md).

The design goal is _MS Office-like operation covering every pptx expression the
library can author_. Generated registry checks track API dispatch coverage;
browser tests must separately establish usable editing workflows.

## API dispatch coverage

The library exposes ~440 public functions. The ones a UI must surface as an
**operation** are the _mutating_ (state-changing) exports — every `add*`,
`set*`, `clear*`, `remove*`, `insert*`, … There are **150** of them today.

That set is the registry coverage target:

1. **`manifest/generate.mjs`** reads the library source, enumerates the mutating
   exports by verb prefix, parses each signature into an operand + parameter
   schema, and writes **`manifest/capabilities.generated.json`** — 150 entries.
2. **`core/registry.ts`** turns _every_ manifest entry into a runnable Command
   that dispatches to the real library function by name (`pptx[id](operand,
…args)`). No stubs: a command is bound to an actual callable or it fails.
3. **`test/editor-capability-coverage.test.ts`** (in the library's own vitest
   suite) independently re-derives the mutating-export set from the compiled
   library and asserts it equals the manifest exactly, and that every id is a
   real callable. If someone adds a new `setX` authoring function, **`pnpm
test` fails** until it is manifested — and therefore wired into the editor.
4. **`test/editor-command-smoke.test.ts`** drives the registry end-to-end
   (author a shape → fill → move → save → reload) to prove the wiring executes,
   not just type-checks.

These checks detect missing registry entries. They do not prove that every
operation has a complete, accessible user interface. `site/test` tests the
rune-backed history model, and `packages/dev/test/browser` exercises the bundled
editor and saved PPTX content.

## How a capability reaches the user

Every capability is reachable by at least one path, in increasing ergonomics:

- **Command palette** (`Ctrl/Cmd+K`) — searchable list of all 150, always
  available. The guaranteed floor.
- **Properties panel** — auto-generated from the manifest: given the current
  selection it lists _every_ capability that can act on it, grouped by category.
  Exhaustive by construction.
- **Ribbon** (`ribbon/config.ts`) — a PowerPoint-style tab/group layout over the
  common commands, with contextual tabs (Shape Format, Table) that appear with
  the matching selection. Ergonomics for the common path, not the coverage
  surface.
- **Direct manipulation** (`canvas/SlideCanvas.svelte`) — the shape moves for
  real on every frame (`applyLive` re-renders the slide via the preview renderer,
  ~6ms; the whole gesture is one undo step committed on release), with:
  - **smart-guide snapping** (`canvas/snapping.ts`) — edges/centres snap to other
    shapes and the slide, drawing pink guide lines;
  - **multi-select** via marquee (rubber-band on empty canvas) and Shift-click,
    with group move;
  - handles to resize, a top handle to rotate (Shift = 15° steps), double-click
    to edit text;
  - **keyboard**: arrow-nudge (Shift = coarse), Ctrl+D duplicate, Ctrl+C/X/V
    copy·cut·paste, Ctrl+A select-all, Delete, Ctrl+±/0 zoom;
  - **zoom / fit** (auto-fit to the viewport, manual zoom in the status bar);
  - a **right-click context menu** (`ui/ContextMenu.svelte`).

  These all funnel through the controller's actions and the same undoable
  command path, so direct manipulation and the ribbon never diverge.

## Argument collection

Commands that need arguments open a dialog (`ui/CommandDialog.svelte`) built from
the parameter schema by `ui/ParamField.svelte`, which renders a control per
kind (string / number / EMU-with-units / color / boolean / enum) and **recurses
into nested object and array schemas** (see the Gradient / Transition dialogs).
Capabilities whose options are enriched with a schema get a field-based form;
the rest fall back to a structured-JSON editor, so the long tail stays usable
while remaining reachable. Enriched schemas come from two places, merged in
`manifest/overrides.ts` (hand wins per id):

- **`manifest/overrides.generated.ts`** — schemas produced by a one-off pass
  that read the library's option types (exact enum members, nested object/array
  fields for `TextFormat`, `TableCellBorders`, `ArrowOptions`, …). Rebuild with
  `manifest/build-generated-overrides.mjs`, which **validates that each
  override's top-level parameter names match the generated capability** — a
  mismatch would make the registry pass the wrong positional args, so it is
  rejected rather than emitted.
- **hand entries in `overrides.ts`** — flagship dialogs (gradient, shadow,
  transition, …) and label/ribbon tuning.

## State & undo

`core/document.svelte.ts` holds the live `PresentationData` in `$state.raw` and
drives re-render with a `version` counter (the library mutates its object graph
in place; deep-proxying fights that). Undo/redo snapshots by **serializing to
`.pptx` bytes** (`savePresentation`/`loadPresentation`) — the library's
guaranteed round-trip — because the model's real state hangs off a symbol-keyed
`OpcPackage` that `structuredClone` silently corrupts. Edits stay synchronous;
snapshots are taken asynchronously (one per discrete gesture).

## Regenerating the manifest

```
node site/src/lib/editor/manifest/generate.mjs
```

Run this whenever the library's authoring surface changes; the coverage test
tells you when it is needed.

## Development preview integration

The editor is currently a separate site route. The completion requirements and
remaining integration work are tracked in [the editor preview roadmap](../../../../docs/editor-preview-roadmap.md).
Command discovery is not proof of complete editing workflows; browser interaction
and persistence checks are required for each supported workflow.
