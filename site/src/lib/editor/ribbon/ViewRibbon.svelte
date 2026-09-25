<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const guides = $derived(editor.guidesVisible());
  function toggleGrid() { editor.view.save({ grid: !editor.view.grid, smart: editor.view.smart, drawing: editor.view.drawing }); }
  function toggleGuides() { editor.view.save({ grid: editor.view.grid, smart: editor.view.smart, drawing: !guides }); }
</script>

<div class="group">
  <div class="items">
    <button aria-pressed={editor.viewMode === 'normal'} onclick={() => { editor.setViewMode('normal'); editor.thumbnailsVisible = true; }}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="1"/><path d="M8 4v16M2 9h6M2 14h6"/></svg>{t('Normal')}</button>
    <button aria-pressed={editor.viewMode === 'sorter'} onclick={() => editor.setViewMode('sorter')}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2 4h8v6H2zM14 4h8v6h-8zM2 14h8v6H2zM14 14h8v6h-8z"/></svg>{t('Slide Sorter')}</button>
  </div>
  <span>{t('Presentation views')}</span>
</div>
<div class="group">
  <div class="checks">
    <label><input type="checkbox" checked={editor.view.ruler} disabled={editor.viewMode !== 'normal'} onchange={() => editor.view.save({ ruler: !editor.view.ruler })} />{t('Ruler')}</label>
    <label><input type="checkbox" checked={editor.view.grid} disabled={editor.viewMode !== 'normal'} onchange={toggleGrid} />{t('Gridlines')}</label>
    <label><input type="checkbox" checked={guides} disabled={editor.viewMode !== 'normal'} onchange={toggleGuides} />{t('Guides')}</label>
    <label><input type="checkbox" bind:checked={editor.thumbnailsVisible} disabled={editor.viewMode !== 'normal'} />{t('Thumbnails')}</label>
    <button class="options" aria-pressed={editor.viewMode === 'normal' && editor.notesVisible} disabled={editor.viewMode !== 'normal'} onclick={() => editor.notesVisible = !editor.notesVisible}>{t('Notes')}</button>
    <button class="options" disabled={editor.viewMode !== 'normal'} onclick={() => editor.activeDialog = 'gridOptions'}>{t('Grid Options...')}</button>
  </div>
  <span>{t('Show')}</span>
</div>
<div class="group">
  <div class="items">
    <button onclick={() => editor.activeDialog = 'zoom'}><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10" cy="10" r="6"/><path d="m15 15 7 7"/></svg>{t('Zoom')}</button>
    <button onclick={() => editor.zoomFit()}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 2H2v6M16 2h6v6M2 16v6h6M22 16v6h-6"/><rect x="6" y="7" width="12" height="10"/></svg>{t('Fit to Window')}</button>
  </div>
  <span>{t('Zoom')}</span>
</div>

<style>
  .group { display: flex; flex-direction: column; justify-content: space-between; flex-shrink: 0; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  .items { display: flex; align-items: center; flex: 1; gap: 2px; }
  button { font: inherit; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  .items button { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 4px; width: 65px; min-height: 55px; font-size: 11px; padding: 4px; }
  button:disabled { opacity: 0.45; cursor: default; }
  label:has(input:disabled) { opacity: 0.45; }
  button:not(:disabled):hover, button[aria-pressed='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  svg { width: 24px; height: 24px; stroke: currentColor; fill: none; stroke-width: 1.2; }
  .group > span { text-align: center; font-size: 10px; color: var(--ok-text-2); padding: 4px 0 1px; }
  .checks { display: grid; grid-template-rows: repeat(3, auto); grid-auto-flow: column; align-content: center; gap: 3px 12px; flex: 1; font-size: 11px; }
  label { display: flex; align-items: center; gap: 4px; }
  input { margin: 0; accent-color: var(--ok-accent); }
  .options { text-align: left; padding: 2px 0; }
</style>
