<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const guides = $derived(editor.guidesVisible());
  function toggleGrid() { editor.view.save({ grid: !editor.view.grid, smart: editor.view.smart, drawing: editor.view.drawing }); }
  function toggleGuides() { editor.view.save({ grid: editor.view.grid, smart: editor.view.smart, drawing: !guides }); }
</script>

<!-- Mac PowerPoint's View tab without the views and masters this editor lacks
     (Notes Page, Reading View, Slide/Handout/Notes Master, Macros). The
     Thumbnails switch and Grid Options live in the View menu, as they do there. -->
<div class="group">
  <button aria-pressed={editor.viewMode === 'normal'} onclick={() => { editor.setViewMode('normal'); editor.thumbnailsVisible = true; }}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="1"/><path d="M8 4v16M2 9h6M2 14h6"/></svg>{t('Normal')}</button>
  <button aria-pressed={editor.viewMode === 'outline'} onclick={() => editor.setViewMode('outline')}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="1"/><path d="M9 4v16M4 8h3M4 11h3M4 14h3M4 17h3"/></svg>{t('Outline View')}</button>
  <button aria-pressed={editor.viewMode === 'sorter'} onclick={() => editor.setViewMode('sorter')}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2 4h8v6H2zM14 4h8v6h-8zM2 14h8v6H2zM14 14h8v6h-8z"/></svg>{t('Slide Sorter')}</button>
</div>
<div class="group">
  <div class="checks">
    <label><input type="checkbox" checked={editor.view.ruler} disabled={editor.viewMode === 'sorter'} onchange={() => editor.view.save({ ruler: !editor.view.ruler })} />{t('Ruler')}</label>
    <label><input type="checkbox" checked={editor.view.grid} disabled={editor.viewMode === 'sorter'} onchange={toggleGrid} />{t('Gridlines')}</label>
    <label><input type="checkbox" checked={guides} disabled={editor.viewMode === 'sorter'} onchange={toggleGuides} />{t('Guides')}</label>
  </div>
  <button aria-pressed={editor.viewMode !== 'sorter' && editor.notesVisible} disabled={editor.viewMode === 'sorter'} onclick={() => editor.notesVisible = !editor.notesVisible}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="2" width="16" height="20" rx="1"/><path d="M8 7h8M8 11h8M8 15h8M8 19h5"/></svg>{t('Notes')}</button>
</div>
<div class="group">
  <button onclick={() => editor.activeDialog = 'zoom'}><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10" cy="10" r="6"/><path d="m15 15 7 7"/></svg>{t('Zoom')}</button>
  <button onclick={() => editor.zoomFit()}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 2H2v6M16 2h6v6M2 16v6h6M22 16v6h-6"/><rect x="6" y="7" width="12" height="10"/></svg>{t('Fit to Window')}</button>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  button { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 4px; width: 58px; min-height: 66px; padding: 4px 2px; font: inherit; font-size: 11px; line-height: 1.2; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:disabled { opacity: 0.45; cursor: default; }
  label:has(input:disabled) { opacity: 0.45; }
  button:not(:disabled):hover, button[aria-pressed='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  svg { width: 32px; height: 32px; flex-shrink: 0; stroke: currentColor; fill: none; stroke-width: 1.1; }
  .checks { display: flex; flex-direction: column; gap: 6px; padding-right: 8px; font-size: 12px; }
  label { display: flex; align-items: center; gap: 6px; }
  input { margin: 0; accent-color: var(--ok-accent); }
</style>
