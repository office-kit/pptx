<script lang="ts">
  // Right-click menu. Items adapt to the current selection and dispatch through
  // the controller's actions (which go through the same undoable command path).
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const menu = $derived(editor.contextMenu!);

  interface Item {
    label: string;
    accel?: string;
    run: () => void;
    disabled?: boolean;
    sep?: boolean;
  }

  const hasShapes = $derived(doc.selection.kind === 'shape' || doc.selection.kind === 'cell');

  const items = $derived.by<Item[]>(() => {
    const list: Item[] = [];
    if (hasShapes) {
      list.push(
        { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
        { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Duplicate', accel: '⌘D', run: () => editor.duplicateSelection() },
        { label: 'Delete', accel: 'Del', run: () => editor.deleteSelection(), sep: true },
        { label: 'Bring to front', run: () => editor.invoke('bringShapeToFront') },
        { label: 'Bring forward', run: () => editor.invoke('bringShapeForward') },
        { label: 'Send backward', run: () => editor.invoke('sendShapeBackward') },
        { label: 'Send to back', run: () => editor.invoke('sendShapeToBack'), sep: true },
        { label: 'Group', run: () => editor.invoke('groupShapes'), disabled: doc.selection.kind !== 'shape' || doc.selection.shapeIds.length < 2 },
        { label: 'Ungroup', run: () => editor.invoke('ungroupShapes') },
      );
    } else {
      list.push(
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Select all', accel: '⌘A', run: () => editor.selectAllShapes() },
      );
    }
    return list;
  });

  function activate(item: Item) {
    if (item.disabled) return;
    item.run();
    editor.closeContextMenu();
  }
</script>

<svelte:window
  onpointerdown={() => editor.closeContextMenu()}
  onblur={() => editor.closeContextMenu()}
/>

<div
  class="ctx ok-editor"
  style="left:{menu.x}px; top:{menu.y}px;"
  role="menu"
  tabindex="-1"
  onpointerdown={(e) => e.stopPropagation()}
  oncontextmenu={(e) => e.preventDefault()}
>
  {#each items as item (item.label)}
    <button class="ctx-item" class:sep={item.sep} role="menuitem" disabled={item.disabled} onclick={() => activate(item)}>
      <span>{t(item.label)}</span>
      {#if item.accel}<span class="accel">{item.accel}</span>{/if}
    </button>
  {/each}
</div>

<style>
  .ctx {
    position: fixed;
    z-index: 400;
    min-width: 200px;
    background: var(--ok-panel);
    border: 1px solid var(--ok-border);
    border-radius: var(--ok-radius-lg);
    box-shadow: var(--ok-shadow-lg);
    padding: 5px;
  }
  .ctx-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 24px;
    width: 100%;
    text-align: left;
    padding: 6px 10px;
    border: none;
    background: none;
    border-radius: var(--ok-radius);
    font: inherit;
    font-size: 13px;
    color: var(--ok-text);
    cursor: pointer;
  }
  .ctx-item:hover:not(:disabled) {
    background: var(--ok-selected);
  }
  .ctx-item:disabled {
    color: var(--ok-text-3);
    cursor: default;
  }
  .ctx-item.sep {
    margin-bottom: 5px;
    padding-bottom: 9px;
    border-bottom: 1px solid var(--ok-border);
  }
  .accel {
    font-size: 11px;
    color: var(--ok-text-3);
  }
</style>
