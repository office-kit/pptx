# @office-kit/pptx-editor — architecture

A ribbon-style presentation editing UI built **entirely on the `@office-kit/pptx` public
API**, in Svelte 5. Hosts embed it with `mountEditor` (`index.ts`; see the
[package README](../README.md)); the docs site serves it that way at `/editor`.
`@office-kit/pptx-dev` bundles it into its local development preview through the
unstable `internal` entry (`internal.ts`, `dev/`), whose host persists edits and
resolves source conflicts; see [the development-tool README](../../dev/README.md).

`mountEditor` renders into a shadow root. Code that listens on `window` or
`document`, reads the focused element or the text selection, or queries the DOM
must go through `core/dom-root.ts` and `EditorController.shell` so it sees the
nodes inside that root; window-level shortcuts check `EditorController.ownsEvent`
so they leave the host page and other editors alone.

The design goal is _familiar desktop-style operation covering every pptx expression the
library can author_. Generated registry checks track API dispatch coverage;
browser tests must separately establish usable editing workflows.

## API dispatch coverage

The library exposes ~440 public functions. The ones a UI must surface as an
**operation** are the _mutating_ (state-changing) exports — every `add*`,
`set*`, `clear*`, `remove*`, `insert*`, … The live count is the `count` field
of `manifest/capabilities.generated.json`; it moves whenever the API grows, so
it is deliberately not repeated here.

That set is the registry coverage target:

1. **`manifest/generate.mjs`** reads the library source, enumerates the mutating
   exports by verb prefix, parses each signature into an operand + parameter
   schema, and writes **`manifest/capabilities.generated.json`** — one entry per
   mutating export.
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
operation has a complete, accessible user interface. `packages/editor/test` tests the
rune-backed history model, `packages/editor/test/browser` embeds the built package
in a plain page, and `packages/dev/test/browser` exercises the bundled
editor and saved PPTX content.

## How a capability reaches the user

Every capability is reachable by at least one path, in increasing ergonomics:

- **Command palette** (`⌘?`, Help ▸ PowerPoint Help, or the title bar's Search) —
  searchable list of every capability, always available. The guaranteed floor.
- **Properties panel** — auto-generated from the manifest: given the current
  selection it lists _every_ capability that can act on it, grouped by category.
  Exhaustive by construction.
- **Menu bar** (`ui/MenuBar.svelte`) — Mac PowerPoint's File … Help menus in the
  title bar, generated from a native capture (`core/menubar-native.ts`); each item
  runs an existing command (`ui/menubar-commands.ts`) and its PowerPoint keyboard
  shortcut runs the same command (`core/menubar-shortcuts.ts`).
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
    with group move, proportional corner resize and a shared rotation handle;
  - handles to resize in the object’s rotated axes (Shift preserves aspect ratio),
    a top handle to rotate (Shift = 15° steps), double-click
    to edit text;
  - **keyboard**: arrow-nudge (Shift = coarse), Delete, and PowerPoint's menu
    shortcuts (⌘D duplicate, ⌘C/X/V, ⌘A, ⌘+/⌘- zoom …; Ctrl works for ⌘ unless
    PowerPoint gives the Control chord its own command);
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
node packages/editor/src/manifest/generate.mjs
```

Run this whenever the library's authoring surface changes; the coverage test
tells you when it is needed.

## Development preview integration

The same editor is available as a site route and bundled in the development
preview. Completion requirements and remaining work are tracked in [the editor preview roadmap](../../../docs/editor-preview-roadmap.md).
Command discovery is not proof of complete editing workflows; browser interaction
and persistence checks are required for each supported workflow.

## Trademarks

Microsoft and PowerPoint are trademarks of the Microsoft group of companies.
This editor is an independent implementation, not affiliated with or endorsed
by Microsoft. It ships no Microsoft icons, artwork, fonts or text; product
names appear only to describe `.pptx` compatibility.
