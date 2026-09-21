<script lang="ts">
  // Right-click menu. Items adapt to the current selection and dispatch through
  // the controller's actions (which go through the same undoable command path).
  import { getEditor } from '../core/context.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const selected = $derived(selectedSlideIndices(doc.selection));
  const firstSelected = $derived(selected[0] ?? 0);
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
    if (doc.selection.kind === 'cell') {
      list.push(
        { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
        { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Clear cell text', accel: 'Del', run: () => editor.deleteSelection(), sep: true },
        { label: 'Select all cells', accel: '⌘A', run: () => editor.selectAll() },
        { label: 'Select table', run: () => {
          const selection = doc.selection;
          if (selection.kind === 'cell') doc.select({ kind: 'shape', slideIndex: selection.slideIndex, shapeIds: [selection.shapeId] });
        } },
      );
    } else if (hasShapes) {
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
        { label: 'Group', run: () => editor.invoke('groupShapes'), disabled: !editor.canRun('groupShapes') },
        { label: 'Ungroup', run: () => editor.invoke('ungroupShapes'), disabled: !editor.canRun('ungroupShapes') },
      );
    } else if (doc.selection.kind === 'slide') {
      list.push(
        { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
        { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'New slide', run: () => editor.invoke('addBlankSlide') },
        { label: 'Duplicate slide', accel: '⌘D', run: () => editor.invoke('duplicateSlide') },
        { label: 'Delete slide', accel: 'Del', run: () => editor.invoke('removeSlide'), sep: true },
        { label: 'Move slide up', run: () => editor.invoke('moveSlide', { toIndex: firstSelected - 1 }), disabled: firstSelected === 0 },
        { label: 'Move slide down', run: () => editor.invoke('moveSlide', { toIndex: firstSelected + 1 }), disabled: firstSelected >= doc.slides.length - selected.length },
      );
    } else {
      list.push(
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Select all', accel: '⌘A', run: () => editor.selectAllShapes() },
      );
    }
    return list;
  });

  let origin: HTMLElement | null = null;
  let menuNode: HTMLDivElement | null = null;
  function dismiss() {
    editor.closeContextMenu();
    if (origin?.isConnected) origin.focus({ preventScroll: true });
  }
  function onKeydown(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); dismiss(); return; }
    if (event.key === 'Tab') { dismiss(); return; }
    if (!menuNode || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const buttons = [...menuNode.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    if (!buttons.length) return;
    const current = buttons.findIndex(button => button === document.activeElement);
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 :
      (current + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length;
    buttons[next]?.focus();
  }

  function place(node: HTMLDivElement, position: { x: number; y: number }) {
    origin = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    menuNode = node;
    node.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus({ preventScroll: true });
    let current = position;
    const update = () => {
      node.style.left = `${Math.max(8, Math.min(current.x, window.innerWidth - node.offsetWidth - 8))}px`;
      node.style.top = `${Math.max(8, Math.min(current.y, window.innerHeight - node.offsetHeight - 8))}px`;
    };
    update();
    window.addEventListener('resize', update);
    return {
      update(position: { x: number; y: number }) { current = position; update(); },
      destroy() { window.removeEventListener('resize', update); },
    };
  }

  function activate(item: Item) {
    if (item.disabled) return;
    dismiss();
    item.run();
  }
</script>

<svelte:window
  onpointerdown={() => editor.closeContextMenu()}
  onblur={() => editor.closeContextMenu()}
/>

<div
  class="ctx ok-editor"
  use:place={menu}
  role="menu"
  tabindex="-1"
  onkeydown={onKeydown}
  onpointerdown={(e) => e.stopPropagation()}
  oncontextmenu={(e) => e.preventDefault()}
>
  {#each items as item (item.label)}
    <button class="ctx-item" class:sep={item.sep} role="menuitem" tabindex="-1" disabled={item.disabled} onclick={() => activate(item)}>
      <span>{t(item.label)}</span>
      {#if item.accel}<span class="accel">{item.accel}</span>{/if}
    </button>
  {/each}
</div>

<style>
  .ctx {
    position: fixed;
    z-index: 400;
    min-width: min(200px, calc(100vw - 16px));
    max-width: calc(100vw - 16px);
    max-height: calc(100dvh - 16px);
    box-sizing: border-box;
    overflow: auto;
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
  .ctx-item:hover:not(:disabled), .ctx-item:focus-visible {
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
