<script lang="ts">
  // Mac PowerPoint's Edit menu. Find and Replace live here rather than on the
  // Home ribbon, as in the native app.
  import { tick } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let root: HTMLDivElement;
  let trigger: HTMLButtonElement;
  const canCopy = $derived(doc.selection.kind === 'slide' || selectedShapeIds(doc.selection).length > 0);
  function close(restore = true) { open = false; if (restore) trigger?.focus(); }
  function choose(action: () => void) { close(); action(); }
  async function show() {
    open = !open;
    if (open) { await tick(); root.querySelector<HTMLButtonElement>('[role="menu"] button:not(:disabled)')?.focus(); }
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const buttons = [...root.querySelectorAll<HTMLButtonElement>('[role="menu"] button:not(:disabled)')];
    const index = buttons.indexOf(event.target as HTMLButtonElement);
    buttons[event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (index + (event.key === 'ArrowUp' ? -1 : 1) + buttons.length) % buttons.length]?.focus();
  }
  const items = $derived([
    { label: 'Undo', key: '⌘Z', disabled: !doc.canUndo, run: () => doc.undo() },
    { label: 'Redo', key: '⌘Y', disabled: !doc.canRedo, run: () => doc.redo(), sep: true },
    { label: 'Cut', key: '⌘X', disabled: !canCopy, run: () => editor.cutSelection() },
    { label: 'Copy', key: '⌘C', disabled: !canCopy, run: () => editor.copySelection() },
    { label: 'Paste', key: '⌘V', disabled: false, run: () => void editor.paste() },
    { label: 'Select All', key: '⌘A', disabled: !doc.currentSlide, run: () => editor.selectAll(), sep: true },
    { label: 'Find...', key: '⌘F', disabled: false, run: () => editor.runOrPrompt('replaceTextInPresentation') },
    { label: 'Replace...', key: '⇧⌘H', disabled: false, run: () => editor.runOrPrompt('replaceTextInPresentation') },
  ]);
</script>
<svelte:window onpointerdown={event => { if (open && !root.contains(event.target as Node)) close(false); }} />
<div class="edit-menu" bind:this={root}>
  <button class="ok-btn" bind:this={trigger} aria-haspopup="menu" aria-expanded={open} onclick={show}>{t('Edit')}</button>
  {#if open}
    <div class="menu" role="menu" tabindex="-1" aria-label={t('Edit')} onkeydown={keys}>
      {#each items as item (item.label)}
        <button role="menuitem" aria-label={t(item.label)} disabled={item.disabled} onclick={() => choose(item.run)}>{t(item.label)}<kbd>{item.key}</kbd></button>
        {#if item.sep}<hr />{/if}
      {/each}
    </div>
  {/if}
</div>

<style>
  .edit-menu { position: relative; }
  .menu { position: absolute; top: 100%; left: 0; z-index: 120; min-width: 200px; padding: 5px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; align-items: center; gap: 6px; border: 0; border-radius: 4px; background: transparent; color: inherit; width: 100%; padding: 4px 7px; text-align: left; white-space: nowrap; font: inherit; font-size: 13px; }
  .menu button:hover:not(:disabled), .menu button:focus-visible { background: var(--ok-accent); color: white; outline: none; }
  .menu button:disabled { opacity: .45; }
  kbd { margin-left: auto; padding-left: 18px; font: inherit; opacity: .65; }
  hr { border: 0; border-top: 1px solid var(--ok-border); margin: 4px 0; }
</style>
