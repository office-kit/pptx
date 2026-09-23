<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { asColor, setShapeImageOpacity, setShapeImageBrightness, setShapeImageContrast, getShapeStrokeEffective, getShapeStrokeColorResolved, getShapeStrokeDash, setShapeStroke, setShapeStrokeDash, getShapePreset, type PresetShape, getShapeKind, getShapeImageCrop, getShapeImageOpacity, getShapeImageBrightness, getShapeImageContrast, getShapeDescription } from '@office-kit/pptx';

  const editor = getEditor();
  const doc = editor.doc;
  const picture = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape' || sel.shapeIds.length !== 1) return null;
    const shape = doc.shapeById(sel.slideIndex, sel.shapeIds[0]!);
    return shape && getShapeKind(shape) === 'picture' ? shape : null;
  });
  const masks: Array<[PresetShape, string]> = [['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Ellipse'], ['triangle', 'Triangle'], ['diamond', 'Diamond'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'], ['star5', 'Star'], ['heart', 'Heart']];
  const mask = $derived.by(() => { doc.version; return picture ? getShapePreset(picture) : null; });
  function setMask(value: string) {
    const preset = masks.find(([key]) => key === value)?.[0];
    if (picture && preset) editor.invoke('setShapePreset', { preset });
  }
  const border = $derived.by(() => {
    doc.version;
    const stroke = picture ? getShapeStrokeEffective(doc.pres, picture) : null;
    const visible = stroke?.kind === 'solid';
    return { visible, color: picture ? getShapeStrokeColorResolved(doc.pres, picture) ?? '#000000' : '#000000', width: visible ? (stroke.widthEmu ?? 9525) / 12700 : 0, dash: picture ? getShapeStrokeDash(picture) ?? 'solid' : 'solid' };
  });
  function borderStyle(value: string) {
    if (!picture) return;
    if (value === 'none') { editor.invoke('setShapeNoStroke'); return; }
    if (value !== 'solid' && value !== 'dash' && value !== 'dot') return;
    doc.transact(t('Image border style'), () => {
      // The reader widens the color to a string; a value the writer would
      // reject leaves the width alone rather than inventing a color.
      const color = asColor(border.color);
      if (!border.visible && color) setShapeStroke(picture!, { color, widthEmu: 12700 });
      setShapeStrokeDash(picture!, value);
    });
  }
  function borderWidth(input: HTMLInputElement) {
    if (!input.reportValidity()) { input.value = String(border.width); return; }
    if (input.valueAsNumber === 0) editor.invoke('setShapeNoStroke');
    else editor.invoke('setShapeStroke', { options: { ...(border.visible ? {} : { color: border.color }), widthEmu: Math.round(input.valueAsNumber * 12700) } });
  }
  const crop = $derived.by(() => { doc.version; return picture ? getShapeImageCrop(picture) : null; });
  const description = $derived.by(() => { doc.version; return picture ? getShapeDescription(picture) ?? '' : ''; });
  const sides = [['left', 'Crop left (%)'], ['top', 'Crop top (%)'], ['right', 'Crop right (%)'], ['bottom', 'Crop bottom (%)']] as const;
  const effects = $derived.by(() => { doc.version; return picture ? [
    { id: 'setShapeImageOpacity', label: 'Opacity (%)', param: 'opacity', value: (getShapeImageOpacity(picture) ?? 1) * 100, min: 0 },
    { id: 'setShapeImageBrightness', label: 'Brightness (%)', param: 'value', value: (getShapeImageBrightness(picture) ?? 0) * 100, min: -100 },
    { id: 'setShapeImageContrast', label: 'Contrast (%)', param: 'value', value: (getShapeImageContrast(picture) ?? 0) * 100, min: -100 },
  ] : []; });
  const hasAdjustments = $derived(effects.some(effect => effect.value !== (effect.param === 'opacity' ? 100 : 0)));
  function resetAdjustments() {
    if (!picture || !hasAdjustments) return;
    doc.transact(t('Reset image adjustments'), () => {
      setShapeImageOpacity(picture!, null);
      setShapeImageBrightness(picture!, null);
      setShapeImageContrast(picture!, null);
    });
  }
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
    <button class="ok-btn" onclick={() => editor.runOrPrompt('setShapeImageCrop')}>{t('Crop image')}</button>
    <label>{t('Image shape')}<select class="ok-input" aria-label={t('Image shape')} value={mask ?? ''} onchange={e => setMask(e.currentTarget.value)}>
      {#if !masks.some(([key]) => key === mask)}<option value={mask ?? ''}>{t('Custom shape')}</option>{/if}
      {#each masks as [key, label]}<option value={key}>{t(label)}</option>{/each}
    </select></label>
    <div class="fields">
      {#each sides as [side, label]}
        <label>{t(label)}<input class="ok-input" type="number" min="0" max="99.9" step="0.1" required value={Math.round((crop?.[side] ?? 0) * 1000) / 10} onchange={(e) => setCrop(side, e.currentTarget)} /></label>
      {/each}
    </div>
    <button class="ok-btn" disabled={!crop} onclick={() => editor.invoke('setShapeImageCrop', { crop: null })}>{t('Reset crop')}</button>
    <div class="fields">
      <label>{t('Image border color')}<input type="color" value={/^#[0-9a-f]{6}$/i.test(border.color) ? border.color : '#000000'} onchange={e => editor.invoke('setShapeStroke', { options: { color: e.currentTarget.value, widthEmu: Math.round((border.width || 1) * 12700) } })} /></label>
      <label>{t('Image border width (points)')}<input class="ok-input" type="number" min="0" max="1584" step="0.25" required value={Math.round(border.width * 100) / 100} onchange={e => borderWidth(e.currentTarget)} /></label>
    </div>
    <label>{t('Image border style')}<select class="ok-input" aria-label={t('Image border style')} value={border.visible ? border.dash : 'none'} onchange={e => borderStyle(e.currentTarget.value)}>
      <option value="none">{t('No outline')}</option><option value="solid">{t('Solid line')}</option><option value="dash">{t('Dashed line')}</option><option value="dot">{t('Dotted line')}</option>
      {#if !['solid', 'dash', 'dot'].includes(border.dash)}<option value={border.dash}>{t('Custom line')}</option>{/if}
    </select></label>
    {#each effects as effect}
      <label>{t(effect.label)}<input class="ok-input" type="number" min={effect.min} max="100" step="1" required value={Math.round(effect.value)} onchange={(e) => {
        if (e.currentTarget.reportValidity()) editor.invoke(effect.id, { [effect.param]: e.currentTarget.valueAsNumber / 100 });
        else e.currentTarget.value = String(Math.round(effect.value));
      }} /></label>
    {/each}
    <button class="ok-btn" disabled={!hasAdjustments} onclick={resetAdjustments}>{t('Reset image adjustments')}</button>
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
