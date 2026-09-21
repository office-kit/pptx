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

For a single-command first launch (macOS/Linux):

```sh
npx --yes @office-kit/pptx-dev@latest init my-slides && cd my-slides && npm install && npm run dev
```

The built-in `init` is the starter generator: no separate `degit` checkout is
needed, and it selects compatible package versions. For an existing project,
start development with `npm run dev` or `npx office-pptx dev deck.tsx`.
If you prefer the bare `office-pptx dev deck.tsx` command, install the CLI once
with `npm install --global @office-kit/pptx-dev@latest`. The local npm script
continues to use the project's installed version.

To update an existing project, stop its dev server and run inside the project:

```sh
npm install -D @office-kit/pptx-dev@latest
npm run dev
```

Open the local URL printed by the server. Save a slide file, `theme.ts` or `deck.tsx` to rebuild. The viewer
has a vertical thumbnail strip, a large slide canvas and an AI chat panel on the right. Click a thumbnail or use
arrow keys, Page Up/Down, Home/End to navigate. Fit/zoom and Present (Escape to
exit) are viewing controls. Select an area for an AI instruction or use **Edit text**
to save a literal directly. All edits are persisted in TSX.
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

Saved TSX edits reuse unchanged source transforms and slide previews to reduce update
latency. The deck is still evaluated in a fresh worker on every build. Changes to
shared resources, including themes, relationships, charts and images, redraw all
slides. Update time depends on the deck size and any code it runs.

## Agents in the preview

The right panel defaults to **Claude Code**. Click **Start** to open the locally
installed `claude` CLI in an interactive terminal. Install and sign in to Claude
Code first. Use a current version with HTTP `UserPromptSubmit` hook support.
The session runs in the deck entry's directory and uses your normal account,
user/project settings, CLAUDE.md, plugins, skills, MCP servers and hooks.

Type `/model` to choose a model, `/config` to change settings, or `/` to browse
available commands and skills. Bash and other tools are available according to
Claude Code's own configuration, and permission prompts are answered directly in
the panel. The preview does not force `acceptEdits`, restrict the tool list or
bypass permissions. Installation and account setup use the CLI's usual flow.
Markdown responses are displayed by Claude Code's terminal renderer; its native
formatting may differ from a web Markdown page. OS-specific actions such as
opening an external editor still use your local environment.

Each submitted Claude prompt receives the focused slide's **1-based number,
text, slide count, preview revision, deck entry and source dependency paths**,
plus any build error, through a session-scoped `UserPromptSubmit` HTTP hook.
The preview captures context with terminal input before forwarding it to the
CLI. Navigation while Claude is replying does not change that turn's context.
Stale input is rejected with a retry message. If your Claude settings or managed
policy disable hooks, automatic context attachment is unavailable; enable the
preview hook in Claude Code to use it. No project settings file is rewritten.

Use **Shift+Enter** to insert a newline in the embedded Claude prompt. On completion,
the preview checks the actual deck build. Build errors are returned to Claude's
Stop hook or to the Codex conversation for up to three repair attempts per user
request. If repair fails, the error stays visible and the last successful preview
is retained. Stopping the agent cancels further repair attempts.

Both agents receive guidance to check text against its actual background, including
inherited theme colors, and to use native `Text` paragraphs or mapped rows for lists
and agendas. `Bullets` and `TableOfContents` are not built-in DSL components; define
or import project helpers if using those names. See the
[TSX authoring reference](../../skill/references/tsx.md#lists-and-tables-of-contents)
for a complete agenda example. Contrast guidance still requires visual review;
a successful build is not a visual accessibility check.
Dependency paths are source candidates, not an exact slide-to-file mapping.
“This slide” requests a focused patch; requests about another slide or the
whole deck can edit other files or shared styling.

Reloading the browser reconnects to the same process and replays its recent
terminal output (up to 2 MB). Each pane has its own session and terminal controls. Keep the owning tab open
while agents are working. **End session** stops the process; saved edits remain. Start opens
a new session, and Claude's `/resume` command can restore an earlier conversation.
Stopping the dev server also stops all embedded sessions. This is independent of
the Claude Code conversation that may have launched your dev server.

Select **Codex** while Claude Code in that pane is stopped to use the existing message-based
chat. Codex replies render Markdown (including lists, code blocks and tables);
raw HTML, unsafe links and remote images are not rendered. Codex still uses
`exec --json` with the `workspace-write` sandbox and its installed CLI login.
Operations requiring interactive approval are not supported in this adapter.
The most recent 12 messages (bounded text) and the focused slide are sent with
each request. The dev server retains up to 100 messages until it exits.
**New chat** clears this history; **Stop** cancels the running edit. Ctrl/Cmd+Enter
sends a message. Model/settings menus in the embedded terminal are currently
specific to Claude Code.

Use the **Split right** or **Split down** icons in a pane's title bar to run up to
four agents side by side or stacked. Each pane independently selects Claude Code
or Codex, retains its own conversation and receives the currently selected slide.
Drag the divider to adjust the split, or focus it and use the arrow keys. Layout
and Claude terminal sessions survive a browser reload. Splitting preserves running
sessions and unsent drafts. **Close** ends only that pane's agent; closing the last
pane opens a fresh one. All agents edit the same project directory, so assign
separate slides or files when working simultaneously.

Drag the left edge of the chat panel to adjust its width. The preview remembers the
width in this browser; double-click the edge to reset it. You can also focus the
edge and use the left/right arrow keys (Shift for larger steps), or Home/End for
the minimum/maximum width. Narrow screens keep the stacked layout.

Use **Agents** in the header to hide/show the panel. On narrow screens the panel
moves below the preview. Presentation mode hides it without ending the session.
Both providers save TSX directly, which triggers the usual automatic rebuild.
The server binds to loopback and requires same-origin JSON for terminal control;
the local context hook uses a separate per-session token.

The terminal uses `node-pty` and xterm.js. Linux systems without a matching PTY
prebuild need the native build prerequisites documented by
[node-pty](https://github.com/microsoft/node-pty). With pnpm, allow `node-pty`'s
install script (`pnpm approve-builds`) if your project blocks dependency scripts.

References: [Claude Code commands](https://code.claude.com/docs/en/commands),
[Claude Code hooks](https://code.claude.com/docs/en/hooks), and
[Codex non-interactive mode](https://developers.openai.com/codex/noninteractive).

## Edit from the slide

Use **Select area** and drag a rectangle, then enter an instruction such as
“Move this down a little.” Choose an agent pane and click **Apply with AI**.
The request includes the slide, preview revision, relative bounds and intersecting
text. Start Claude Code in that pane first; Codex starts when you send.
Selection is cleared when the preview or selected slide changes. Shift+Enter
inserts a newline; Cmd/Ctrl+Enter sends the request.

Use **Edit text**, click a paragraph, and **Save text** to update its TSX source
without waiting for an agent. Direct saves require a unique source literal within
the entry directory and verify that other slides did not change. Computed text,
shared values and ambiguous matches use **Apply with AI** instead. **Undo text**
restores the last direct edit only if its file has not changed since. **Undo with
AI** asks the agent to reverse its previous selection edit while preserving newer
work; it is not a source snapshot restore.

After each successful agent edit or direct text save/undo, the dev server captures
all changed slides, including slides offscreen. Codex receives PNG attachments;
Claude Code receives local PNG paths through its Stop hook and is instructed to
read them. The agent checks overlap, clipping, alignment, spacing and contrast,
and can make up to two correction passes followed by a final inspection. This is
AI-assisted review, not a guarantee of pixel-perfect rendering. The preview updates
as soon as the build finishes, independently of screenshot review.

Screenshot capture uses an installed Chrome, or falls back to Playwright Chromium.
Set `PLAYWRIGHT_CHANNEL` to use another installed Chromium channel. If neither is
available, install Chrome and restart the dev server. Capture failures are shown
explicitly; saved edits remain available. If the chosen agent is busy or not started,
use **Retry visual review** after making it available. Screenshots and review prompts
are stored locally under `.office-kit/reviews/`; new projects ignore this directory.
Add `.office-kit/` to `.gitignore` in existing projects. Images are provided to the
selected AI agent using its existing login and provider settings.

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
