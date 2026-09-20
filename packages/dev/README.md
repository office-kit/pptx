# @office-kit/pptx-dev

Local TSX authoring tools for `@office-kit/pptx-dsl`. Build and preview an editable
PowerPoint presentation while changing its source in VSCode or Claude Code.
Requires Node.js 22.18 or later.

The CLI and DSL packages are **not yet published to npm**. Start with the
[step-by-step authoring guide](https://office-kit.github.io/pptx/docs/authoring)
for a source-checkout installation. In the repository root (after
`pnpm install --frozen-lockfile`), run:

```sh
pnpm --filter @office-kit/pptx-dev... build
PPTX_REPO="$PWD"
PPTX_PACKS="$PPTX_REPO/.probe/packages"
mkdir -p "$PPTX_PACKS"
pnpm --filter @office-kit/pptx-dev... exec pnpm pack --pack-destination "$PPTX_PACKS"
node "$PPTX_REPO/packages/dev/dist/cli.mjs" init ../my-slides
cd ../my-slides
npm install --legacy-peer-deps "$PPTX_PACKS"/*.tgz
npm run check
npm run dev
```

Install all four tarballs from the same checkout. The temporary peer override is
needed because the DSL targets the next core release, whose version number has
not been bumped yet; the source checkout contains the required core changes.
Keep the tarballs for reinstalls, and use an empty pack directory after upgrading
so old versions are not included. The guide also covers Claude Code, VSCode,
templates, viewing controls, and troubleshooting.

Open the local URL printed by the server. Save `deck.tsx` to rebuild. The viewer
has a vertical thumbnail strip and a large slide canvas. Click a thumbnail or use
arrow keys, Page Up/Down, Home/End to navigate. Fit/zoom and Present (Escape to
exit) are viewing controls; the canvas has no editing, dragging or resize handles.
Changes are made only in TSX, including when an AI agent edits the presentation.
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
workflow. Preview-to-source selection is not implemented.

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
A fresh worker evaluates each build, with a 60-second limit, so module state
cannot accumulate across updates. Workers are not a security sandbox. The
preview server binds to `127.0.0.1` and rejects unexpected Host headers.

Programmatic exports: `buildDeck`, `exportDeck`, `initProject`, `inspectTemplate`.
`buildDeck` returns PPTX bytes, SVG slides, aspect ratio, module dependencies and
core validation diagnostics. Browser APIs and DOM globals are not available in
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
