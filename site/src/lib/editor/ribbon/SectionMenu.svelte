<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import {
    addSection,
    removeAllSections,
    removeSection,
    renameSection,
    sectionOf,
    sectionRanges,
    UNTITLED_SECTION,
  } from '../core/sections.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';

  let { small = false }: { small?: boolean } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const ranges = $derived.by(() => { doc.version; return sectionRanges(doc.pres); });
  const current = $derived(sectionOf(ranges, doc.selection.slideIndex));
  let open = $state(false);
  let renaming = $state<{ start: number; name: string } | null>(null);
  let dialog = $state<HTMLDialogElement>();
  $effect(() => { if (renaming && dialog && !dialog.open) dialog.showModal(); });

  function add() {
    open = false;
    const start = doc.selection.slideIndex;
    doc.transact(t('Add Section'), () => addSection(doc.pres, start, t(UNTITLED_SECTION)));
    renaming = { start, name: t(UNTITLED_SECTION) };
  }
  function rename(event: SubmitEvent) {
    event.preventDefault();
    if (!renaming) return;
    const { start, name } = renaming;
    renaming = null;
    doc.transact(t('Rename Section'), () => renameSection(doc.pres, start, name));
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !(event.target as Element).closest?.('.section-menu')) open = false; }} />

<div class="section-menu">
  <button class:big={!small} class:row={small} aria-label={t('Section')} aria-haspopup="menu" aria-expanded={open} disabled={!doc.currentSlide} onclick={() => (open = !open)}>
    {#if small}<Icon name="section" size={18} /><span>{t('Section')}</span><span aria-hidden="true">⌄</span>{:else}<span class="icon-row"><Icon name="section" size={32} /><span aria-hidden="true">⌄</span></span><span>{t('Section')}</span>{/if}
  </button>
  {#if open}
    <div class="menu" role="menu" aria-label={t('Section')}>
      <button role="menuitem" onclick={add}>{t('Add Section')}</button>
      <button role="menuitem" disabled={!current} onclick={() => { open = false; if (current) renaming = { start: current.start, name: current.name }; }}>{t('Rename Section')}</button>
      <button role="menuitem" disabled={!current} onclick={() => { open = false; if (current) { const start = current.start; doc.transact(t('Remove Section'), () => removeSection(doc.pres, start)); } }}>{t('Remove Section')}</button>
      <hr />
      <button role="menuitem" disabled={ranges.length === 0} onclick={() => { open = false; doc.transact(t('Remove All Sections'), () => removeAllSections(doc.pres)); }}>{t('Remove All Sections')}</button>
    </div>
  {/if}
</div>
{#if renaming}
  <dialog bind:this={dialog} aria-label={t('Rename Section')} onclose={() => (renaming = null)}>
    <form onsubmit={rename}>
      <label>{t('Section name:')}<input class="ok-input" bind:value={renaming.name} /></label>
      <footer><button type="button" class="ok-btn" onclick={() => (renaming = null)}>{t('Cancel')}</button><button type="submit" class="ok-btn primary">{t('Rename')}</button></footer>
    </form>
  </dialog>
{/if}

<style>
  .section-menu { position: relative; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  .big { display: flex; flex-direction: column; align-items: center; gap: 2px; min-width: 44px; padding: 3px 3px; font-size: 11px; line-height: 1.15; }
  .icon-row { display: flex; align-items: center; gap: 2px; }
  .row { display: flex; align-items: center; gap: 4px; padding: 2px 4px; font-size: 11px; white-space: nowrap; }
  .menu { position: absolute; top: 100%; left: 0; z-index: 400; display: flex; flex-direction: column; min-width: 200px; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { padding: 5px 8px; text-align: left; font-size: 12px; }
  hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 4px 0; }
  dialog { width: min(340px, 90vw); padding: 16px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); }
  dialog::backdrop { background: #0006; }
  form { display: grid; gap: 12px; }
  label { display: grid; gap: 6px; font-size: 12px; }
  footer { display: flex; justify-content: flex-end; gap: 8px; }
</style>
