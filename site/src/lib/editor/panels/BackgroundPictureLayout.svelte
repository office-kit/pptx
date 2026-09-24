<script lang="ts">
  import { getSlides, getSlidePartName, getSlideBackgroundImageFillLayout, setSlideBackgroundImageFillLayout, pt, type ImageFillLayout, type ImageTileAlignment, type ImageTileFlip, type SlideData } from '@office-kit/pptx';
  import { switchRememberedImageLayout } from '../core/remembered-image-fill.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const { doc } = getEditor();
  const slides = $derived.by(() => { doc.version; const all = getSlides(doc.pres); return selectedSlideIndices(doc.selection).flatMap(index => all[index] ? [all[index]!] : []); });
  const layouts = $derived(slides.map(getSlideBackgroundImageFillLayout));
  const tiled = $derived(layouts.length > 0 && layouts.every(layout => layout?.mode === 'tile'));
  const mixedMode = $derived(!tiled && layouts.some(layout => layout?.mode === 'tile'));
  const alignments: [ImageTileAlignment, string][] = [['tl', 'Top left'], ['t', 'Top'], ['tr', 'Top right'], ['l', 'Left'], ['ctr', 'Center'], ['r', 'Right'], ['bl', 'Bottom left'], ['b', 'Bottom'], ['br', 'Bottom right']];
  const flips: [ImageTileFlip, string][] = [['none', 'None'], ['x', 'Horizontal'], ['y', 'Vertical'], ['xy', 'Both']];
  const tileFields = [['offsetX', 'Offset X'], ['offsetY', 'Offset Y'], ['scaleX', 'Scale X'], ['scaleY', 'Scale Y']] as const;
  const stretchFields = [['left', 'Offset left'], ['right', 'Offset right'], ['top', 'Offset top'], ['bottom', 'Offset bottom']] as const;
  let error = $state('');
  function common<T>(values: T[]): T | undefined { return values.every(value => value === values[0]) ? values[0] : undefined; }
  const alignment = $derived(common(layouts.map(value => value?.mode === 'tile' ? value.alignment ?? 'tl' : undefined)));
  const flip = $derived(common(layouts.map(value => value?.mode === 'tile' ? value.flip ?? 'none' : undefined)));
  function edit(change: (layout: ImageFillLayout, slide: SlideData) => ImageFillLayout) {
    try { doc.transact(t('Picture or texture fill'), () => { for (const slide of slides) { const layout = getSlideBackgroundImageFillLayout(slide); if (layout) setSlideBackgroundImageFillLayout(slide, change(layout, slide)); } }); error = ''; }
    catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  function switchLayout(tile: boolean) {
    edit((current, slide) => {
      const key = `background:${getSlidePartName(slide)}`;
      const remembered = doc.rememberedFills.get(key) ?? {};
      const layouts = remembered.imageLayouts ??= {};
      const next = switchRememberedImageLayout(current, tile ? 'tile' : 'stretch', layouts);
      doc.rememberedFills.set(key, remembered);
      return next;
    });
  }
  function numericValue(field: string) {
    return common(layouts.map(value => {
      if (value?.mode === 'tile') {
        if (field === 'offsetX' || field === 'offsetY') return (value[field] ?? 0) / pt(1);
        if (field === 'scaleX' || field === 'scaleY') return (value[field] ?? 1) * 100;
      } else if (value?.mode === 'stretch' && (field === 'left' || field === 'right' || field === 'top' || field === 'bottom')) return (value[field] ?? 0) * 100;
      return undefined;
    }));
  }
  function numeric(input: HTMLInputElement, field: string) {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { input.value = String(numericValue(field) ?? ''); return; }
    const value = input.valueAsNumber;
    edit(current => ({ ...current, [field]: field.startsWith('offset') ? pt(value) : value / 100 }));
  }
</script>

<div class="picture-layout">
  <label class="check"><input type="checkbox" checked={tiled} indeterminate={mixedMode} onchange={event => switchLayout(event.currentTarget.checked)} />{t('Tile picture as texture')}</label>
  {#each tiled ? tileFields : stretchFields as [field, label]}
    <label class="row"><span>{t(label)}</span><span class="number"><input class="ok-input" type="number" min={field.startsWith('scale') ? 0 : tiled ? -1584 : -100000} max={field.startsWith('scale') ? 100 : tiled ? 1584 : 100000} step="any" value={numericValue(field) ?? ''} placeholder={t('Mixed')} aria-label={t(label)} onchange={event => numeric(event.currentTarget, field)} />{field.startsWith('offset') ? 'pt' : '%'}</span></label>
  {/each}
  {#if tiled}
    <label class="row">{t('Alignment')}<select class="ok-input" aria-label={t('Alignment')} value={alignment ?? ''} onchange={event => { const value = alignments.find(([key]) => key === event.currentTarget.value)?.[0]; if (value) edit(current => current.mode === 'tile' ? { ...current, alignment: value } : current); }}>
      {#if alignment === undefined}<option value="" disabled>{t('Mixed')}</option>{/if}
      {#each alignments as [value, label]}<option {value}>{t(label)}</option>{/each}
    </select></label>
    <label class="row">{t('Mirror type')}<select class="ok-input" aria-label={t('Mirror type')} value={flip ?? ''} onchange={event => { const value = flips.find(([key]) => key === event.currentTarget.value)?.[0]; if (value) edit(current => current.mode === 'tile' ? { ...current, flip: value } : current); }}>
      {#if flip === undefined}<option value="" disabled>{t('Mixed')}</option>{/if}
      {#each flips as [value, label]}<option {value}>{t(label)}</option>{/each}
    </select></label>
  {/if}
  <label class="check"><input type="checkbox" disabled />{t('Rotate with shape')}</label>
  {#if error}<p role="alert">{error}</p>{/if}
</div>

<style>
  .picture-layout { display: grid; gap: 8px; font-size: 11px; }
  .row, .check, .number { display: flex; align-items: center; gap: 6px; }
  .row { justify-content: space-between; }
  .row select { max-width: 130px; }
  .number input { width: 64px; min-width: 0; }
  [role=alert] { color: #bf3131; }
</style>
