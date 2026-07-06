<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { getShapeKind, getShapeName } from '@office-kit/pptx';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

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
  <div class="zoom">
    <button class="zbtn" title={t('Zoom out (Ctrl+-)')} onclick={() => editor.zoomOut()}>−</button>
    <button class="zpct" title={t('Reset to 100%')} onclick={() => editor.zoomReset()}>{Math.round(editor.zoom * 100)}%</button>
    <button class="zbtn" title={t('Zoom in (Ctrl+=)')} onclick={() => editor.zoomIn()}>+</button>
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
  .zoom button {
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
