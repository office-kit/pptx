<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getTableCell, getTableCellParagraphs, setTableCellClickAction, getShapeParagraphCount, getShapeParagraphElements, getShapeRunClickAction, type ShapeClickAction, getShapeRunHyperlinkTooltip, getShapeClickAction, getSlideIndex, getSlideTitle, setShapeClickAction, getShapeHyperlink, getShapeHyperlinkTooltip, getShapeKind, getShapeText, setShapeHyperlink } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const range = untrack(() => editor.linkTextRange);
  const cellPosition = untrack(() => editor.linkTableCell);
  const selection = untrack(() => doc.selection);
  const version = untrack(() => doc.version);
  const shapes = selectedShapeIds(selection).map(id => doc.shapeById(selection.slideIndex, id));
  const slides = untrack(() => doc.slides);
  const supported = selection.kind === 'shape' && shapes.length > 0 && shapes.every(shape => shape && ['shape', 'picture', 'connector', 'graphicFrame'].includes(getShapeKind(shape)));
  const selectedRuns = shapes.flatMap(shape => {
    if (!shape || !range) return [];
    if (cellPosition) {
      let offset = 0;
      return getTableCellParagraphs(getTableCell(shape, cellPosition.row, cellPosition.col)).flatMap(paragraph => {
        const runs = paragraph.elements.flatMap(element => {
          const length = element.kind === 'br' ? 1 : element.text.length;
          const selected = offset < range.end && offset + length > range.start;
          offset += length;
          return selected ? [{ action: element.clickAction ?? null, tip: element.tooltip ?? null }] : [];
        });
        offset++;
        return runs;
      });
    }
    const runs: { action: ShapeClickAction | null; tip: string | null }[] = [];
    let offset = 0;
    for (let p = 0; p < getShapeParagraphCount(shape); p++) {
      let r = 0;
      for (const element of getShapeParagraphElements(shape, p)) {
        const length = element.kind === 'br' ? 1 : element.text.length;
        if (offset < range.end && offset + length > range.start) {
          runs.push(element.kind === 'r' ? { action: getShapeRunClickAction(shape, p, r), tip: getShapeRunHyperlinkTooltip(shape, p, r) } : { action: null, tip: null });
        }
        if (element.kind === 'r') r++;
        offset += length;
      }
      offset++;
    }
    return runs;
  });
  const actions = range ? selectedRuns.map(run => run.action) : shapes.map(shape => {
    if (!shape) return null;
    const url = getShapeHyperlink(shape);
    return url ? { kind: 'url' as const, url } : getShapeClickAction(shape);
  });
  const keys = actions.map(action => action?.kind === 'url' ? `url:${action.url}` : action?.kind === 'slide' ? `slide:${getSlideIndex(doc.pres, action.slide)}` : action?.kind ?? '');
  const mixed = keys.some(key => key !== keys[0]);
  const initial = mixed ? null : actions[0];
  const presets = ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide'] as const;
  let destination = $state(initial?.kind ?? 'url');
  let slideIndex = $state(initial?.kind === 'slide' ? getSlideIndex(doc.pres, initial.slide) : selection.slideIndex);
  let url = $state(initial?.kind === 'url' ? initial.url : '');
  const tips = range ? selectedRuns.map(run => run.tip) : shapes.map(shape => shape ? getShapeHyperlinkTooltip(shape) : null);
  let tooltip = $state(tips.every(tip => tip === tips[0]) ? tips[0] ?? '' : '');
  const valid = $derived(destination === 'slide' ? !!slides[slideIndex] : destination === 'url' ? !!url.trim() : presets.some(kind => kind === destination));
  let error = $state('');
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  function apply(remove = false) {
    if (!supported || (!remove && !valid)) return;
    if (doc.version !== version || doc.selection !== selection) { error = t('The selection changed. Reopen this dialog.'); return; }
    try {
      doc.transact(t(remove ? 'Remove link' : 'Edit link'), () => {
        for (const shape of shapes) if (shape) {
          if (range && cellPosition) {
            setTableCellClickAction(getTableCell(shape, cellPosition.row, cellPosition.col), remove ? null : destination === 'url' ? { kind: 'url', url: url.trim() } : destination === 'slide' ? { kind: 'slide', slide: slides[slideIndex]! } : { kind: destination }, { range, tooltip: tooltip.trim() || undefined });
            continue;
          }
          if (range) {
            if (remove) setShapeClickAction(shape, null, { range });
            else if (destination === 'url') setShapeHyperlink(shape, url.trim(), tooltip.trim() || undefined, { range });
            else setShapeClickAction(shape, destination === 'slide' ? { kind: 'slide', slide: slides[slideIndex]! } : { kind: destination }, { range, tooltip: tooltip.trim() || undefined });
            continue;
          }
          const hasText = getShapeKind(shape) === 'shape' && getShapeText(shape).length > 0;
          if (hasText) setShapeHyperlink(shape, null);
          setShapeClickAction(shape, null);
          if (!remove) {
            if (destination === 'slide') setShapeClickAction(shape, { kind: 'slide', slide: slides[slideIndex]! }, { tooltip: tooltip.trim() || undefined });
            else if (destination !== 'url') setShapeClickAction(shape, { kind: destination }, { tooltip: tooltip.trim() || undefined });
            else if (hasText) setShapeHyperlink(shape, url.trim(), tooltip.trim() || undefined);
            else setShapeClickAction(shape, { kind: 'url', url: url.trim() }, { tooltip: tooltip.trim() || undefined });
          }
        }
      });
      editor.closeDialog();
    } catch (cause) { error = `${t('Link update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`; }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Edit link')} onclose={() => editor.closeDialog()}>
  <form onsubmit={(event) => { event.preventDefault(); apply(); }}>
    <header><strong>{t('Edit link')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    {#if supported}
      <p>{t(range ? 'Applies to the selected text.' : 'Applies to the selected objects.')}</p>
      <label>{t('Link destination')}<select class="ok-input" bind:value={destination} aria-label={t('Link destination')}><option value="url">{t('Web address')}</option><option value="slide">{t('Slide in this presentation')}</option><option value="nextSlide">{t('Next slide')}</option><option value="prevSlide">{t('Previous slide')}</option><option value="firstSlide">{t('First slide')}</option><option value="lastSlide">{t('Last slide')}</option></select></label>
      {#if destination === 'url'}
      <label>{t('Link address')}<input class="ok-input" type="url" required bind:value={url} placeholder="https://example.com" aria-label={t('Link address')} /></label>
      {:else if destination === 'slide'}
        <label>{t('Target slide')}<select class="ok-input" bind:value={slideIndex} aria-label={t('Target slide')}>{#each slides as slide, i}<option value={i}>{i + 1}. {getSlideTitle(slide) || t('Untitled slide')}</option>{/each}</select></label>
      {/if}
      {#if mixed}<p>{t(range ? 'The selected text contains different links.' : 'The selected shapes have different links.')}</p>{/if}
      <label>{t('Link description')}<input class="ok-input" bind:value={tooltip} aria-label={t('Link description')} /></label>
    {:else}<p role="alert">{t('Select objects to edit their links.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" disabled={!supported || !actions.some(Boolean)} onclick={() => apply(true)}>{t('Remove link')}</button><span></span><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!supported || !valid}>{t('Apply')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(520px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form, label { display: grid; gap: 12px; }
  header, footer { display: flex; align-items: center; gap: 10px; }
  header { justify-content: space-between; }
  footer span { flex: 1; }
  label, p { font-size: 12px; }
  p { color: var(--ok-text-2); margin: 0; }
  [role='alert'] { color: #bf3131; }
</style>
