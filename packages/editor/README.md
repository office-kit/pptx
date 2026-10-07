# @office-kit/pptx-editor

A browser editor for PowerPoint presentations that you embed in your own web
application. It opens a `.pptx`, lets people edit it with a ribbon, menus,
slide thumbnails and direct manipulation, and hands the edited `.pptx` back.
Everything it does goes through the public API of
[`@office-kit/pptx`](https://github.com/office-kit/pptx), so the saved file is
an ordinary, editable presentation.

> **Experimental (0.x).** The editor's behaviour is still growing and a minor
> release may change the API. Pin a version.

```sh
npm install @office-kit/pptx-editor
```

## Embedding

```ts
import { mountEditor } from '@office-kit/pptx-editor';

const editor = mountEditor(document.getElementById('editor')!, {
  source: new Uint8Array(await (await fetch('/decks/q3.pptx')).arrayBuffer()),
  locale: 'en',
  onSave: async (pptx) => {
    await fetch('/decks/q3.pptx', { method: 'PUT', body: pptx });
  },
});

await editor.ready; // the deck is open
const pptx = await editor.snapshot(); // the deck as it is now
editor.destroy(); // remove the editor again
```

The editor fills the element you mount it in, so give that element a size
(for example `height: 100vh`).

### `mountEditor(target, options?)`

| Option   | Type                                          | Default                                                     |
| -------- | --------------------------------------------- | ----------------------------------------------------------- |
| `source` | `Uint8Array`                                  | A new presentation with one title slide, as File ▸ New.     |
| `locale` | `'en' \| 'ja'`                                | The language last picked in the editor, else the browser's. |
| `onSave` | `(pptx: Uint8Array) => Promise<void> \| void` | Without it, saving downloads the file.                      |

`onSave` runs when the user saves: the Save button, File ▸ Save, or
Ctrl/Cmd+S. It receives the complete `.pptx`. If it throws or rejects, the
editor tells the user the save failed.

It returns an `EditorHandle`:

| Member               | Description                                                                                                      |
| -------------------- | ---------------------------------------------------------------------------------------------------------------- |
| `ready`              | Resolves once `source` is open and the editor is shown; rejects if `source` is not a readable `.pptx`.           |
| `snapshot()`         | The presentation as it is now, as `.pptx` bytes. It does not call `onSave` or mark the deck saved.               |
| `selection()`        | The shapes the user has selected, as `ShapeRef`s (see [Agents in the browser](#agents-in-the-browser)).          |
| `apply(label, edit)` | Runs `edit(presentation)` with the `@office-kit/pptx` API as one undo step named "Agent: `label`".               |
| `tools()`            | Every editing capability as tool definitions for a language model (see [Tools for a model](#tools-for-a-model)). |
| `run(name, input)`   | Runs one of those tools with the model's JSON input, as one undo step named "Agent: `name`".                     |
| `on(type, listener)` | Calls `listener` on `'change'` (`{ source: 'user' \| 'agent' }`) or `'selectionchange'`; returns `off()`.        |
| `destroy()`          | Removes the editor and every listener it added, leaving `target` as it was, ready to mount again.                |

`'change'` fires after every edit the presentation keeps: the user's, an
`apply`, Undo, Redo, and File ▸ New / Open. It does not fire for the source
you passed, nor while a shape is still being dragged.

### React

```tsx
import { useEffect, useRef } from 'react';
import { mountEditor } from '@office-kit/pptx-editor';

export function DeckEditor({
  source,
  onSave,
}: {
  source: Uint8Array;
  onSave: (pptx: Uint8Array) => Promise<void>;
}) {
  const target = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const editor = mountEditor(target.current!, { source, onSave });
    return () => editor.destroy();
  }, [source, onSave]);
  return <div ref={target} style={{ height: '100vh' }} />;
}
```

### Vue

```vue
<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref } from 'vue';
import { mountEditor, type EditorHandle } from '@office-kit/pptx-editor';

const props = defineProps<{ source: Uint8Array; onSave: (pptx: Uint8Array) => Promise<void> }>();
const target = ref<HTMLElement>();
let editor: EditorHandle | undefined;
onMounted(() => {
  editor = mountEditor(target.value!, { source: props.source, onSave: props.onSave });
});
onBeforeUnmount(() => editor?.destroy());
</script>

<template>
  <div ref="target" style="height: 100vh" />
</template>
```

## Agents in the browser

An agent running in your page — an LLM with tools, a script, a chat panel next
to the editor — works on the same document the user sees, through the handle:

- **Read**: `snapshot()` gives the `.pptx`; open it with `loadPresentation`
  from `@office-kit/pptx` to read text, layout and anything else. To _look_ at
  a slide, render it with `renderSlideToSvg` from `@office-kit/pptx-preview`,
  the renderer the editor itself uses.
- **"This shape"**: `selection()` returns what the user selected as
  `ShapeRef`s — `{ slideIndex, slide, shapeId, name }`. A ref names the slide
  by its part name and the shape by its id, so it stays valid while slides
  move and across Undo. `resolveShape(presentation, ref)` finds the shape again
  and throws if it was deleted, so an agent never edits the wrong shape.
- **Write**: `apply(label, edit)` runs `edit` against the live presentation
  with the `@office-kit/pptx` API. The user sees the change at once, and it is
  one undo step named "Agent: `label`" (「エージェント: `label`」 in Japanese).
  It marks the deck as changed exactly as a user edit does; `onSave` still runs
  only when the user saves. If `edit` throws, nothing of it is kept and
  `apply` rejects with the error. `edit` must be synchronous.
- **Follow along**: `on('change', …)` reports every kept edit with its
  `source` (`'user'` or `'agent'`), including Undo and Redo;
  `on('selectionchange', …)` reports the user's new selection.

`apply` is for code you write. When a model decides the edits, use
[`tools()` and `run`](#tools-for-a-model) instead: they check the model's
input for you.

```ts
import { mountEditor, resolveShape } from '@office-kit/pptx-editor';
import { setShapeFill } from '@office-kit/pptx';

const editor = mountEditor(target, { source });
await editor.ready;

const [selected] = editor.selection();
if (selected)
  await editor.apply('Highlight', (presentation) => {
    setShapeFill(resolveShape(presentation, selected), '#FFF2CC');
  });

editor.on('change', ({ source }) => {
  if (source === 'user') {
    // The user edited (or undid an agent's edit): refresh what the agent knows.
  }
});
```

There is no `renderSlide` on the handle: `snapshot()` plus
`@office-kit/pptx-preview` already does it, with the preview's own options.

### Tools for a model

`tools()` describes every mutating `@office-kit/pptx` export as a tool a
language model can call, and `run(name, input)` runs the call the model makes.
There is no per-function glue to write, and a new library function becomes a
tool with the next release.

- **Names** are the export names (`setShapeFill`, `addSlideShape`, …), sorted.
  Four tools are the editor's own, for finding what the others take:
  `listSlides`, `listShapes` (refs, kind, bounds and text of each shape),
  `listLayouts` and `listComments`.
- **Descriptions** are the exports' TSDoc. **Input schemas** are JSON Schema
  (draft 2020-12) generated from the functions' TypeScript signatures, one
  property per parameter. The open presentation is never an input; a slide is
  its part name (`"/ppt/slides/slide1.xml"`), a shape a `ShapeRef`, a table
  cell a `ShapeRef` plus `row` and `col`, a layout its part name. Lengths are
  integers in EMU (914400 per inch, 12700 per point). Bytes (pictures, media)
  are base64. Each schema is self-contained (`$defs`, `$ref`).
- **`run`** checks `input` against the schema before anything happens and
  never coerces: `"5"` is not a number. On a mismatch it rejects with every
  problem by its JSON Pointer, for example
  `Invalid input for setShapePosition:\n- /x: expected integer, got string "1in"`;
  send that back to the model as the tool's error and it can correct the call.
  A ref to a slide or shape that is gone also rejects, editing nothing. An
  editing tool is one undo step, named "Agent: `setShapeFill`"; a read tool
  edits nothing. It resolves with the result as JSON: `null` for most tools,
  the ref of the slide or shape a tool added, the count a `clear…` tool reports.
- **Not tools**: `createPresentation`, `importSlide`, `mergePresentations` and
  `sortSlides`, which take another presentation or a function.

Each element of `tools()` is `{ name, description, input_schema }`, which the
Anthropic Messages API takes as it is. For OpenAI's function calling, pass
`{ type: 'function', function: { name, description, parameters: input_schema } }`.

All the tools together are large (on the order of 150k tokens; the chart tools
alone are about 30k). Give the model the ones its job needs:

```ts
import Anthropic from '@anthropic-ai/sdk';
import { mountEditor } from '@office-kit/pptx-editor';

const editor = mountEditor(target, { source });
await editor.ready;

const wanted = new Set([
  'listSlides',
  'listShapes',
  'setShapeText',
  'setShapeFill',
  'addSlideShape',
]);
const tools = (await editor.tools()).filter((tool) => wanted.has(tool.name));

// Keep your API key on your server: point the SDK at a proxy that adds it.
const client = new Anthropic({
  baseURL: '/api/anthropic',
  apiKey: 'unused',
  dangerouslyAllowBrowser: true,
});

async function instruct(instruction: string): Promise<void> {
  const messages: Anthropic.MessageParam[] = [
    {
      role: 'user',
      content: `${instruction}\n\nSelected shapes: ${JSON.stringify(editor.selection())}`,
    },
  ];
  for (;;) {
    const response = await client.messages.create({
      model: 'claude-opus-5-5',
      max_tokens: 16000,
      tools,
      messages,
    });
    messages.push({ role: 'assistant', content: response.content });
    if (response.stop_reason !== 'tool_use') return;
    const results: Anthropic.ToolResultBlockParam[] = [];
    for (const block of response.content) {
      if (block.type !== 'tool_use') continue;
      try {
        const result = await editor.run(block.name, block.input);
        results.push({
          type: 'tool_result',
          tool_use_id: block.id,
          content: JSON.stringify(result),
        });
      } catch (error) {
        results.push({
          type: 'tool_result',
          tool_use_id: block.id,
          content: error instanceof Error ? error.message : String(error),
          is_error: true,
        });
      }
    }
    messages.push({ role: 'user', content: results });
  }
}
```

## Behaviour on your page

- **Styles are isolated.** The editor renders inside a shadow root on an
  element it adds to `target`. Your page's stylesheets do not reach it, and
  inherited text styles (font, line height, letter spacing …) stop at its
  boundary; its styles do not reach your page.
- **Keyboard shortcuts stay in the editor.** Keys typed into your page's own
  inputs and buttons are left alone. Keys typed while nothing on the page has
  focus go to the editor used last, as in a desktop application.
- **Several editors can share a page.** Each has its own presentation, undo
  history and selection. The interface language is page-wide: every editor on
  a page uses the same one.
- **Preferences are remembered in `localStorage`** (the language picked in the
  editor, view options such as rulers and gridlines). A `locale` you pass is not
  stored.
- Slide Show and the Agents pane are only available in the
  [`@office-kit/pptx-dev`](../dev) preview, which hosts the editor in a frame.

## Why the editor ships compiled

The editor is written in [Svelte](https://svelte.dev). The package contains
the compiled JavaScript with the small Svelte runtime bundled in, so your
application needs no Svelte, no build plugin and no particular framework, and a
Svelte version of your own never conflicts with the editor's. `@office-kit/pptx`
and `@office-kit/pptx-preview` stay ordinary dependencies.

## Contributing

How the editor is put together, and how its command registry follows the
library's API, is described in [`src/README.md`](src/README.md).

`@office-kit/pptx-editor/internal` and
`@office-kit/pptx-editor/internal/animation-player` are **not public API**.
They exist only for `@office-kit/pptx-dev` and change or disappear in any
release.

## Trademarks

Microsoft and PowerPoint are trademarks of the Microsoft group of companies.
This editor is an independent project, not affiliated with or endorsed by
Microsoft.
