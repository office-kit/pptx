---
name: office-kit-pptx
description: Create and revise editable PowerPoint presentations with typed TSX and a live slide preview. Use for a new slide deck, changes to an existing PPTX or template, or iterative presentation design with office-kit. Handles project setup, preview, checking and PPTX export.
---

# Create slides with office-kit

Turn the user's brief into an editable PPTX. You operate the authoring tools;
the user describes the content and reviews slides in the browser. Make content
and layout changes in TSX. The preview also supports region instructions and direct
text saves. When it provides screenshots after an edit, inspect every image and
correct clipping, overlap, spacing and contrast on affected slides.

## Start or resume a project

For an existing office-kit slide project, read its `CLAUDE.md`, `package.json`
and entry file to locate the requested slides, then use its installed dependencies
and scripts. Read the relevant slide sources and their dependencies as needed;
do not load every slide for a local revision.
Do not initialize over existing work or upgrade packages just to edit a deck.

For a new deck, choose a new child directory in the user's workspace (for example,
`my-slides`). Run the setup yourself; do not hand the user a list of setup commands.
Prerequisites are Node.js 22.18+ and npm. If either is missing, explain the missing
prerequisite instead of attempting system-wide installation.

```sh
npx --yes --registry=https://registry.npmjs.org @office-kit/pptx-dev@latest init my-slides
cd my-slides
npm install --registry=https://registry.npmjs.org
```

Use the published packages. No clone of the office-kit repository, pnpm build,
local tarballs or peer-dependency overrides are needed. If setup fails, report
the actual error and resolve it before proceeding. The initializer refuses an
existing directory; reuse a presentation project or choose a new directory.
Read the generated `CLAUDE.md` before authoring. It is not necessarily loaded
automatically when Claude was started in the parent directory.

## Author and revise

Read [the TSX reference](references/tsx.md) when first authoring or using unfamiliar
elements; reuse already-read guidance during revisions. It covers
native elements, charts, tables, template selection and the Raw escape hatch.
The default export of `deck.tsx` is a `Presentation`. Use ordinary TypeScript
functions and data to reuse content; do not add React, Vue or a web UI.

- Use the supplied brief, audience and source material. Ask only for information
  needed to proceed; label illustrative data when real figures are unavailable.
- Choose a consistent palette and typography, with clear hierarchy and enough
  space for text. Inspect chart labels, table density, contrast and overflow.
- Keep text, shapes, charts and tables as native objects. Do not rasterize whole
  slides to make a preview look correct.
- Keep assets and imported components under the slide project directory so they
  are watched. Load files relative to their source module with
  `readFile(new URL('./asset.png', import.meta.url))`.
- For an existing PPTX, copy it into the project, inspect it with
  `npx --no-install office-pptx inspect template.pptx`, then use source bytes and
  `mode="edit"`. Preserve unmentioned slides, masters, layouts and unknown parts.
  Use compose mode only when the request calls for a new slide sequence. Keep
  the source file separate from the exported `deck.pptx`.
- Prefer typed elements. When a needed capability is only in the core library,
  use `Raw` with public core APIs and consult [the core reference](references/core-api.md)
  and installed type declarations. Do not invent props or describe Raw support
  as complete declarative coverage.

## Keep revisions small

For new decks, keep slide order in `deck.tsx`, each slide in a descriptively named
`slides/*.tsx` file, and shared design values in `theme.ts`. Keep filenames stable
when reordering slides. Reuse ordinary functions for shared elements. Existing
projects can keep their structure; do not split or migrate them for a small edit.

For a request such as “shorten slide 3's headline”, find the slide through the entry
file (or search for its visible text), read its source and necessary dependencies,
and patch only that text. For mapped slides, locate the corresponding data item.
Do not rewrite the whole file/deck or reformat unrelated code. Inspect the diff
when available to ensure unrelated content and slide order are preserved.

Change shared theme values or components only when the request applies to their
consumers. For a one-slide exception, override that slide's props instead. For a
shared change, inspect the affected consumers. Do not add a second source of truth
or an editing DSL: TSX and ordinary code patches are the editing interface.

## Preview for the user

Run `npm run dev` in the slide project using your terminal tool's background-task
support. Wait for `Preview: ...`, open that exact URL if browser tools are
available, and give it to the user. Keep the server running while iterating and
reuse it on follow-up edits. If the port is busy, use `npm run dev -- --port 0`
and read the assigned URL. Do not stop unrelated processes.

Saving TSX updates the preview. The user selects slides in the vertical thumbnail
strip, uses Fit/zoom, or chooses Present for presentation mode. Review affected slides
with available browser/image tools and fix problems in TSX. Keep this loop to
source edits and the running preview; do not run a separate export or restart
the server for each intermediate change. If visual inspection
is unavailable, state that limitation rather than claiming a visual check.

A failed build leaves the last successful preview visible with an error. Fix the
error before treating the visible slides or Download PPTX as current output.
The viewer does not contain an AI chat; the conversation stays in Claude Code.

## Check and deliver

Run both commands after the final edits:

```sh
npm run check
npm run build
```

Type checking is separate from the live rebuild. Fix errors and review all slides,
including template slides retained in edit mode. The SVG preview is a rendering
aid, not a guarantee of PowerPoint fidelity or animation/media playback.

Deliver the preview URL, the generated `deck.pptx` path and the editable project
path. State any unverified rendering or unsupported requested feature. On the
next revision, patch the same project and use the running preview. Rerun the
checks and export when delivering the revised output.
