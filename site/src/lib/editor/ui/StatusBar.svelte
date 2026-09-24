<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { getShapeKind, getShapeName } from '@office-kit/pptx';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const zoomPercent = $derived(Math.round((editor.viewMode === 'sorter' ? editor.sorterZoom : editor.zoom) * 100));
  const sliderPosition = $derived(zoomPercent <= 100 ? (zoomPercent - 10) / 90 * 1000 : 1000 + (zoomPercent - 100) / 300 * 1000);
  function slideZoom(event: Event) {
    const position = Number((event.currentTarget as HTMLInputElement).value);
    editor.setZoom(Math.round(position <= 1000 ? 10 + position / 1000 * 90 : 100 + (position - 1000) / 1000 * 300) / 100);
  }
  function zoomKeys(event: KeyboardEvent) {
    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    const amount = event.shiftKey ? 10 : 1;
    editor.setZoom((event.key === 'Home' ? 10 : event.key === 'End' ? 400 : zoomPercent + (['ArrowLeft', 'ArrowDown'].includes(event.key) ? -amount : amount)) / 100);
  }

  const selectionLabel = $derived.by(() => {
    const sel = doc.selection;
    if (sel.kind === 'shape') {
      const id = selectedShapeId(sel);
      const shape = id == null ? null : doc.shapeById(sel.slideIndex, id);
      if (shape) {
        const extra = sel.shapeIds.length > 1 ? ` +${sel.shapeIds.length - 1}` : '';
        try {
          return `${getShapeName(shape) || getShapeKind(shape)}${extra}`;
        } catch {
          return `Shape${extra}`;
        }
      }
    }
    if (sel.kind === 'cell') return `${t('Cell')} (${sel.row + 1}, ${sel.col + 1})`;
    return t('No selection');
  });
</script>

<div class="statusbar">
  <span>{t('Slide')} {doc.selection.slideIndex + 1} / {doc.slides.length}</span>
  <span class="sep"></span>
  <span>{selectionLabel}</span>
  <span class="spacer"></span>
  <button class="notes-toggle" aria-pressed={editor.notesVisible && editor.viewMode === 'normal'} onclick={() => { if (editor.viewMode === 'normal' && editor.notesVisible) editor.notesVisible = false; else editor.showNotes(); }}>{t('Notes')}</button>
  <div class="views" role="group" aria-label={t('Presentation views')}>
    <button title={t('Normal')} aria-label={t('Normal')} aria-pressed={editor.viewMode === 'normal'} onclick={() => editor.setViewMode('normal')}><svg width="16" height="12" viewBox="0 0 16 12" aria-hidden="true"><rect x=".5" y=".5" width="15" height="11" rx="1"/><path d="M4 1v10M1 4h3M1 8h3"/></svg></button>
    <button title={t('Slide Sorter')} aria-label={t('Slide Sorter')} aria-pressed={editor.viewMode === 'sorter'} onclick={() => editor.setViewMode('sorter')}><svg width="16" height="12" viewBox="0 0 16 12" aria-hidden="true"><path d="M1 1h5v4H1zM9 1h5v4H9zM1 7h5v4H1zM9 7h5v4H9z"/></svg></button>
  </div>
  <div class="zoom">
    <button class="zbtn" title={t('Zoom out (Ctrl+-)')} onclick={() => editor.zoomOut()}>−</button>
    <input type="range" min="0" max="2000" step="1" value={sliderPosition} aria-label={t('Zoom percentage')} aria-valuetext="{zoomPercent}%" oninput={slideZoom} onkeydown={zoomKeys} />
    <button class="zbtn" title={t('Zoom in (Ctrl+=)')} onclick={() => editor.zoomIn()}>+</button>
    <button class="zpct" title={t('Zoom...')} onclick={() => editor.activeDialog = 'zoom'}>{zoomPercent}%</button>
    <button class="zfit" title={t('Fit (Ctrl+0)')} onclick={() => editor.zoomFit()}>{t('Fit')}</button>
  </div>
</div>

<style>
  .statusbar {
    display: flex;
    align-items: center;
    gap: 10px;
    height: 26px;
    padding: 0 12px;
    background: var(--ok-accent);
    color: #fff;
    font-size: 11px;
  }
  .sep {
    width: 1px;
    height: 14px;
    background: rgba(255, 255, 255, 0.3);
  }
  .spacer {
    flex: 1;
  }
  .zoom {
    display: flex;
    align-items: center;
    gap: 2px;
  }
  .zoom input[type='range'] { width: 110px; height: 12px; accent-color: white; margin: 0 4px; }
  .views { display: flex; gap: 3px; }
  .views svg { fill: none; stroke: currentColor; }
  .views button[aria-pressed="true"] { background: rgba(255, 255, 255, .25); }
  .notes-toggle, .views button, .zoom button {
    background: transparent;
    border: none;
    color: #fff;
    font: inherit;
    font-size: 12px;
    cursor: pointer;
    padding: 2px 6px;
    border-radius: 3px;
  }
  .zoom button:hover {
    background: rgba(255, 255, 255, 0.18);
  }
  .zbtn {
    font-size: 15px !important;
    line-height: 1;
    width: 22px;
  }
  .zpct {
    min-width: 46px;
    text-align: center;
  }
  .zfit {
    border: 1px solid rgba(255, 255, 255, 0.4) !important;
    margin-left: 4px;
  }
</style>
