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
const pptx = await editor.save(); // the deck as it is now
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

| Member      | Description                                                                                            |
| ----------- | ------------------------------------------------------------------------------------------------------ |
| `ready`     | Resolves once `source` is open and the editor is shown; rejects if `source` is not a readable `.pptx`. |
| `save()`    | The presentation as it is now, as `.pptx` bytes. It does not call `onSave`.                            |
| `destroy()` | Removes the editor and every listener it added, leaving `target` as it was, ready to mount again.      |

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
