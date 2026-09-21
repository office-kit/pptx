<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { getShapeKind, getShapeImageCrop, getShapeImageOpacity, getShapeImageBrightness, getShapeImageContrast, getShapeDescription } from '@office-kit/pptx';

  const editor = getEditor();
  const doc = editor.doc;
  const picture = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape' || sel.shapeIds.length !== 1) return null;
    const shape = doc.shapeById(sel.slideIndex, sel.shapeIds[0]!);
    return shape && getShapeKind(shape) === 'picture' ? shape : null;
  });
  const crop = $derived.by(() => { doc.version; return picture ? getShapeImageCrop(picture) : null; });
  const description = $derived.by(() => { doc.version; return picture ? getShapeDescription(picture) ?? '' : ''; });
  const sides = [['left', 'Crop left (%)'], ['top', 'Crop top (%)'], ['right', 'Crop right (%)'], ['bottom', 'Crop bottom (%)']] as const;
  const effects = $derived.by(() => { doc.version; return picture ? [
    { id: 'setShapeImageOpacity', label: 'Opacity (%)', param: 'opacity', value: (getShapeImageOpacity(picture) ?? 1) * 100, min: 0 },
    { id: 'setShapeImageBrightness', label: 'Brightness (%)', param: 'value', value: (getShapeImageBrightness(picture) ?? 0) * 100, min: -100 },
    { id: 'setShapeImageContrast', label: 'Contrast (%)', param: 'value', value: (getShapeImageContrast(picture) ?? 0) * 100, min: -100 },
  ] : []; });
  function setCrop(side: 'left' | 'top' | 'right' | 'bottom', input: HTMLInputElement) {
    const next = { ...crop, [side]: input.valueAsNumber / 100 };
    if (!input.reportValidity() || (next.left ?? 0) + (next.right ?? 0) >= 1 || (next.top ?? 0) + (next.bottom ?? 0) >= 1) {
      editor.toast('error', t('Crop must leave part of the image visible'));
      input.value = String((crop?.[side] ?? 0) * 100);
      return;
    }
    editor.invoke('setShapeImageCrop', { crop: next });
  }
</script>

{#if picture}
  <section class="image-controls" aria-label={t('Image options')}>
    <strong>{t('Image options')}</strong>
    <button class="ok-btn" onclick={() => editor.runOrPrompt('setShapeImage')}>{t('Replace image')}</button>
    <div class="fields">
      {#each sides as [side, label]}
        <label>{t(label)}<input class="ok-input" type="number" min="0" max="99.9" step="0.1" required value={Math.round((crop?.[side] ?? 0) * 1000) / 10} onchange={(e) => setCrop(side, e.currentTarget)} /></label>
      {/each}
    </div>
    <button class="ok-btn" disabled={!crop} onclick={() => editor.invoke('setShapeImageCrop', { crop: null })}>{t('Reset crop')}</button>
    {#each effects as effect}
      <label>{t(effect.label)}<input class="ok-input" type="number" min={effect.min} max="100" step="1" required value={Math.round(effect.value)} onchange={(e) => {
        if (e.currentTarget.reportValidity()) editor.invoke(effect.id, { [effect.param]: e.currentTarget.valueAsNumber / 100 });
        else e.currentTarget.value = String(Math.round(effect.value));
      }} /></label>
    {/each}
    <label>{t('Alternative text')}<textarea class="ok-input" rows="3" value={description} onchange={(e) => editor.invoke('setShapeDescription', { description: e.currentTarget.value })}></textarea></label>
  </section>
{/if}

<style>
  .image-controls { padding: 12px; display: flex; flex-direction: column; gap: 10px; border-bottom: 1px solid var(--ok-border); }
  strong { font-size: 12px; }
  label { display: grid; gap: 4px; font-size: 11px; color: var(--ok-text-2); min-width: 0; }
  .fields { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  input, textarea { width: 100%; box-sizing: border-box; }
</style>
