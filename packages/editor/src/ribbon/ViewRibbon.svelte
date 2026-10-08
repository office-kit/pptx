<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import { isSlideEditingView } from '../core/view-modes.ts';
  const editor = getEditor();
  const guides = $derived(editor.guidesVisible());
  function toggleGrid() { editor.view.save({ grid: !editor.view.grid, smart: editor.view.smart, drawing: editor.view.drawing }); }
  function toggleGuides() { editor.view.save({ grid: editor.view.grid, smart: editor.view.smart, drawing: !guides }); }
</script>

<!-- The reference desktop app's (Mac, 16) View tab: five views | three masters | Ruler,
     Gridlines, Guides and Notes | Zoom, Fit to Window | Macros, unchanged
     from 1512 to 1200 pt. The Thumbnails switch and Grid Options live in the
     View menu, as they do there. -->
<div class="view-tab" lang={getLocale()}>
  <section class="cluster" role="group" aria-label={t('Presentation Views')}>
    <button class="big" style:--w="43px" aria-pressed={editor.viewMode === 'normal'} onclick={() => { editor.setViewMode('normal'); editor.thumbnailsVisible = true; }}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="1"/><path d="M8 4v16M2 9h6M2 14h6"/></svg><span>{t('Normal')}</span></button>
    <button class="big" style:--w="43px" aria-pressed={editor.viewMode === 'outline'} onclick={() => editor.setViewMode('outline')}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="1"/><path d="M9 4v16M4 8h3M4 11h3M4 14h3M4 17h3"/></svg><span>{t('Outline View')}</span></button>
    <button class="big" style:--w="38px" aria-pressed={editor.viewMode === 'sorter'} onclick={() => editor.setViewMode('sorter')}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2 4h8v6H2zM14 4h8v6h-8zM2 14h8v6H2zM14 14h8v6h-8z"/></svg><span>{t('Slide Sorter')}</span></button>
    <button class="big" style:--w="38px" aria-pressed={editor.viewMode === 'notesPage'} onclick={() => editor.setViewMode('notesPage')}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="5" y="2" width="14" height="20" rx="1"/><rect x="7" y="4" width="10" height="7"/><path d="M8 14h8M8 17h8M8 20h5"/></svg><span>{t('Notes Page')}</span></button>
    <button class="big" style:--w="47px" onclick={() => editor.openReadingView()}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2 5h9v15H2zM13 5h9v15h-9z"/><path d="M4 9h5M4 12h5M15 9h5M15 12h5"/></svg><span>{t('Reading View')}</span></button>
  </section>
  <section class="cluster" role="group" aria-label={t('Master Views')}>
    <button class="big" style:--w="41px" aria-pressed={editor.viewMode === 'slideMaster'} onclick={() => editor.setViewMode('slideMaster')}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="1"/><path d="M2 8h20M5 12h14M5 15h10"/></svg><span>{t('Slide Master')}</span></button>
    <button class="big" style:--w="49px" aria-pressed={editor.viewMode === 'handoutMaster'} onclick={() => editor.setViewMode('handoutMaster')}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="2" width="16" height="20" rx="1"/><path d="M7 5h4v4H7zM13 5h4v4h-4zM7 11h4v4H7zM13 11h4v4h-4z"/></svg><span>{t('Handout Master')}</span></button>
    <button class="big" style:--w="41px" aria-pressed={editor.viewMode === 'notesMaster'} onclick={() => editor.setViewMode('notesMaster')}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="2" width="16" height="20" rx="1"/><rect x="7" y="4" width="10" height="7"/><path d="M8 14h8M8 17h8"/></svg><span>{t('Notes Master')}</span></button>
  </section>
  <section class="cluster" role="group" aria-label={t('Show')}>
    <div class="checks">
      <label><input type="checkbox" checked={editor.view.ruler} disabled={editor.viewMode === 'sorter'} onchange={() => editor.view.save({ ruler: !editor.view.ruler })} />{t('Ruler')}</label>
      <label><input type="checkbox" checked={editor.view.grid} disabled={editor.viewMode === 'sorter'} onchange={toggleGrid} />{t('Gridlines')}</label>
      <label><input type="checkbox" checked={guides} disabled={editor.viewMode === 'sorter'} onchange={toggleGuides} />{t('Guides')}</label>
    </div>
    <button class="big" style:--w="38px" aria-pressed={isSlideEditingView(editor.viewMode) && editor.notesVisible} disabled={!isSlideEditingView(editor.viewMode)} onclick={() => editor.notesVisible = !editor.notesVisible}><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="2" width="16" height="20" rx="1"/><path d="M8 7h8M8 11h8M8 15h8M8 19h5"/></svg><span>{t('Notes')}</span></button>
  </section>
  <section class="cluster" role="group" aria-label={t('Zoom')}>
    <button class="big" style:--w="38px" onclick={() => editor.activeDialog = 'zoom'}><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10" cy="10" r="6"/><path d="m15 15 7 7"/></svg><span>{t('Zoom')}</span></button>
    <button class="big" style:--w="47px" onclick={() => editor.zoomFit()}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 2H2v6M16 2h6v6M2 16v6h6M22 16v6h-6"/><rect x="6" y="7" width="12" height="10"/></svg><span>{t('Fit to Window')}</span></button>
  </section>
  <section class="cluster" role="group" aria-label={t('Macros')}>
    <button class="big" style:--w="43px" disabled title={t('Macros (VBA) do not run in this editor.')}><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 4l12 8-12 8z"/></svg><span>{t('Macros')}</span></button>
  </section>
</div>

<style>
  /* Geometry measured from the reference desktop app (Mac, 16) (NATIVE_PARITY.md, "Native
     geometry audit"); the Show checkboxes are 22 pt rows on a 19 pt pitch. */
  .view-tab { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
  .cluster { display: flex; flex: none; align-items: stretch; padding: 0 10px; border-right: 1px solid var(--ok-border); }
  .cluster:first-child { padding-left: 4px; }
  .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled, label:has(input:disabled) { opacity: 0.4; cursor: default; }
  button[aria-pressed='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: var(--w); padding: 4px 1px; font-size: 11px; line-height: 1.15; text-align: center; }
  .big > span { max-width: var(--w); margin: 0 -2px; }
  /* Japanese labels wrap per character, so they get at least six characters
     a line and a smaller size that fits three lines (the reference desktop app widens them). */
  .big > span:last-child:lang(ja) { max-width: max(calc(var(--w) - 4px), 6em); font-size: 10px; line-height: 1.1; }
  svg { width: 32px; height: 32px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.1; }
  .checks { display: flex; flex-direction: column; gap: 0; margin-top: 3px; font-size: 12px; }
  .checks label { display: flex; align-items: center; gap: 6px; height: 19px; padding-right: 8px; white-space: nowrap; }
  input { margin: 0; accent-color: var(--ok-accent); }
</style>
