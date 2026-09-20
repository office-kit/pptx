# @office-kit/pptx-dev

Local TSX authoring tools for `@office-kit/pptx-dsl`. Build and preview an editable
PowerPoint presentation while changing its source in VSCode or Claude Code.
Requires Node.js 22.18 or later.

For Claude Code, [install the office-kit skill](https://office-kit.github.io/pptx/docs/authoring)
and ask it to create a presentation. The skill handles setup, preview and export.

To create a project yourself:

```sh
npx --yes @office-kit/pptx-dev@latest init my-slides
cd my-slides
npm install
npm run dev
```

Open the local URL printed by the server. Save a slide file, `theme.ts` or `deck.tsx` to rebuild. The viewer
has a vertical thumbnail strip, a large slide canvas and an AI chat panel on the right. Click a thumbnail or use
arrow keys, Page Up/Down, Home/End to navigate. Fit/zoom and Present (Escape to
exit) are viewing controls; the canvas has no editing, dragging or resize handles.
Changes are made only in TSX, including when an AI agent edits the presentation.
Keep the server running throughout the edit/review loop. Saving updates only the
changed thumbnails and slide view, preserving zoom, scroll position and presentation
mode. The previous slide stays visible until its replacement is ready. Rapid edits
cancel obsolete evaluations; only the latest successful result is published.
Download PPTX exports the last successful build.
A syntax or runtime error is shown without discarding the last successful preview.
DSL evaluation errors include the TSX element's source file and line number.

```sh
npm run check
npm run build
```

`check` runs TypeScript; `build` writes `deck.pptx`. Compilation alone does not
perform TypeScript checking, so run both before delivery. Exported charts, text,
images and tables remain editable. The renderer previews the serialized PPTX,
but is not a guarantee of identical rendering in PowerPoint.

The generated `CLAUDE.md` explains authoring conventions. VSCode picks up the
included `tsconfig.json` for completion and diagnostics. Start the **Preview
presentation** task, then use **Simple Browser: Show** with the printed URL and
move it to a side editor group. No custom editor extension is needed for this
workflow. Clicking a slide supplies chat context; opening its exact TSX location is not implemented.

## Chat in the preview

Choose **Claude Code** or **Codex** in the right panel and describe the change.
Install the selected CLI (`claude` or `codex`) on your PATH and sign in through
that CLI once before using chat. The dev server runs it from the deck entry's
directory, using its existing account and model configuration. No API key is
stored in the browser. Requests use the selected provider and its normal usage
limits. Sending a message authorizes edits to the local project.

Each request includes the focused slide's **1-based number, text, slide count,
preview revision, deck entry and source dependency paths**, plus any build error.
The focus is captured at send time, so navigating during a response does not
change it. A stale preview is rejected with a retry message. Dependency paths are
source candidates, not an exact slide-to-file mapping; the agent locates the
relevant TSX using the entry, slide order and text. This also works for single-file
decks, generated slide lists and imported presentations.

“This slide” directs a focused patch; requests mentioning other slides or the
whole deck can edit other files or shared styling. Saving uses the same automatic
preview rebuild as manual editing. Both CLIs run non-interactively:

- Claude Code uses print mode with JSON events and `acceptEdits`, exposing only
  `Read`, `Edit`, `Write`, `Glob` and `Grep`. Shell execution is not exposed.
- Codex uses `exec --json` with the `workspace-write` sandbox. Sandbox bypass is
  never enabled. Operations requiring interactive approval are not supported.

Progress, replies and failures appear in the panel. **Stop** terminates the running
agent; edits already saved remain. Only one edit runs at a time across browser
tabs. **New chat** clears conversational context without reverting files. Recent
12 messages (bounded text) are passed to each new CLI invocation; these are new
CLI sessions, not a continuation of an existing terminal conversation. The server
keeps up to 100 messages in memory, shared by its tabs, until it exits. Reloading
restores that conversation. Restarting the server clears it.

Use **Chat** in the header to hide/show the panel, and Ctrl/Cmd+Enter to send.
On narrow screens the panel moves below the preview. Presentation mode hides it.
Chat requests require same-origin JSON as well as the server's loopback Host
check. Authentication and CLI installation still happen in a terminal; normal
slide editing and review can then stay in the preview.

Protocol references: [Codex non-interactive mode](https://developers.openai.com/codex/noninteractive)
and [Claude Code programmatic use](https://code.claude.com/docs/en/headless).

## Edit only what changed

New projects separate slide order, content and shared styling:

```text
deck.tsx          # Slide imports and order
slides/cover.tsx  # One slide per descriptively named file
theme.ts          # Shared design values
```

For example, a headline revision changes only the text in the corresponding
slide file. A deck-wide palette change belongs in `theme.ts`; a one-slide color
exception belongs in that slide's props. Imported files are watched, so keep the
preview running while editing. Existing single-file decks remain supported and
do not need migration for a small revision.

Use focused code patches for revisions. Review affected slides while iterating;
check types, export and review the whole deck before delivery. The build still
evaluates the whole presentation; file splitting reduces authoring scope, not
the amount of PPTX evaluation.

## Commands

- `office-pptx init <new-directory>` creates a TSX project without overwriting an existing directory.
- `office-pptx dev <deck.tsx> [--port 4173]` watches the entry file's directory recursively.
- `office-pptx build <deck.tsx> [--out deck.pptx]` writes a PPTX.
- `office-pptx inspect <template.pptx>` prints slides, shape targets and layout references as JSON.

Keep imported components, source PPTX files and images inside the watched
directory. `node_modules`, `.git`, `dist` and `.office-kit` are excluded. Load
assets with `readFile(new URL('./asset.png', import.meta.url))`; the compiler
preserves the original URL of each bundled source module.

TSX runs as trusted local Node code, including imports and Raw callbacks.
The watch server preloads libraries in the next worker between edits. Each build
still evaluates the complete TSX in a fresh worker, with a 60-second limit, so module state
cannot accumulate across updates. Workers are not a security sandbox. The
preview server binds to `127.0.0.1` and rejects unexpected Host headers.

Programmatic exports: `buildDeck`, `exportDeck`, `initProject`, `inspectTemplate`.
`buildDeck` returns PPTX bytes, SVG slides, aspect ratio, module dependencies and
slide text and core validation diagnostics. Browser APIs and DOM globals are not available in
TSX evaluation.

## Developing in this monorepo

The new packages must be built locally before running the CLI; they need not be
published to use this checkout.

```sh
pnpm --filter @office-kit/pptx-dev... build
node packages/dev/dist/cli.mjs dev packages/dsl/examples/review.tsx
node packages/dev/dist/cli.mjs build packages/dsl/examples/review.tsx --out review.pptx
pnpm --filter @office-kit/pptx-dev test
```

Browser regression tests cover partial updates, viewport preservation, errors,
reconnection and presentation mode, and report save-to-visible timings for a
50-slide text deck:

```sh
pnpm --filter @office-kit/pptx-dev exec playwright install chromium
pnpm --filter @office-kit/pptx-dev test:browser
```

To use an installed Chrome instead, set `PLAYWRIGHT_CHANNEL=chrome`. Timings depend
on the deck and machine; complex templates and charts still incur full generation
and serialization costs. SVG updates are transferred as deltas, with a complete
snapshot after reconnecting or falling behind.
