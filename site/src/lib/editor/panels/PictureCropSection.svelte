<script lang="ts">
  import { cm, emu, getShapeBoundsResolved, getShapeImageCrop, setShapeBounds, setShapeImageCrop, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { editFrameCropGeometry, editPictureCropGeometry, getPictureCropGeometry, resetCropGeometry, type PictureCropGeometry } from '../core/crop-geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { picture }: { picture: SlideShapeData } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const maxDimension = 5963.92;
  const locked = $derived(editor.selectionLocked());
  const current = $derived.by(() => {
    doc.version;
    const frame = getShapeBoundsResolved(doc.pres, picture);
    const crop = getShapeImageCrop(picture);
    // Imported OOXML may have an empty source rectangle, which cannot be
    // represented by finite picture dimensions in the numeric crop controls.
    if (!frame || frame.w <= 0 || frame.h <= 0 ||
      1 - (crop?.left ?? 0) - (crop?.right ?? 0) <= 0 ||
      1 - (crop?.top ?? 0) - (crop?.bottom ?? 0) <= 0) return null;
    return { frame, crop, values: getPictureCropGeometry(frame, crop) };
  });
  const fields = [
    { legend: 'Picture position', items: [
      { key: 'pictureWidth', label: 'Width', name: 'Picture width', size: true },
      { key: 'pictureHeight', label: 'Height', name: 'Picture height', size: true },
      { key: 'offsetX', label: 'Offset X', name: 'Offset X', size: false },
      { key: 'offsetY', label: 'Offset Y', name: 'Offset Y', size: false },
    ] },
    { legend: 'Crop position', items: [
      { key: 'cropWidth', label: 'Width', name: 'Crop width', size: true },
      { key: 'cropHeight', label: 'Height', name: 'Crop height', size: true },
      { key: 'cropLeft', label: 'Left', name: 'Crop left', size: false },
      { key: 'cropTop', label: 'Top', name: 'Crop top', size: false },
    ] },
  ] as const;
  const display = (value: number) => Math.round(value / cm(1) * 100) / 100;

  function apply(next: ReturnType<typeof editPictureCropGeometry> | ReturnType<typeof resetCropGeometry>) {
    doc.transact(t('Crop'), () => {
      // Validate crop before mutating bounds: ST_Percentage has a signed limit.
      setShapeImageCrop(picture, next.crop);
      setShapeBounds(picture, { x: emu(next.frame.x), y: emu(next.frame.y), w: emu(next.frame.w), h: emu(next.frame.h) });
    });
  }

  function change(key: keyof PictureCropGeometry, input: HTMLInputElement) {
    const state = current;
    if (!state) return;
    if (locked || !input.reportValidity() || !Number.isFinite(input.valueAsNumber)) {
      input.value = String(display(state.values[key])); return;
    }
    const value = cm(input.valueAsNumber);
    if (display(state.values[key]) === input.valueAsNumber) return;
    const frameKeys = { cropWidth: 'w', cropHeight: 'h', cropLeft: 'x', cropTop: 'y' } as const;
    const next = key in frameKeys
      ? editFrameCropGeometry(state.frame, state.crop, { [frameKeys[key as keyof typeof frameKeys]]: value })
      : editPictureCropGeometry(state.frame, state.crop, { [key]: value });
    try { apply(next); } catch (error) {
      if (!(error instanceof RangeError)) throw error;
      input.value = String(display(state.values[key]));
      editor.toast('error', t('These crop dimensions exceed the supported range.'));
    }
  }

  function reset() {
    if (!current || locked) return;
    apply(resetCropGeometry(current.frame, current.crop));
  }
</script>

<section aria-label={t('Crop')}>
  <details open>
    <summary>{t('Crop')}</summary>
    {#if current}
      {#each fields as group}
        <fieldset disabled={locked}>
          <legend>{t(group.legend)}</legend>
          {#each group.items as field}
            <label><span>{t(field.label)}</span><span class="number"><input class="ok-input" type="number" required aria-label={t(field.name)} min={field.size ? 0.01 : -maxDimension} max={maxDimension} step="0.01" value={display(current.values[field.key])} onchange={e => change(field.key, e.currentTarget)} /><span>cm</span></span></label>
          {/each}
        </fieldset>
      {/each}
      <button class="ok-btn" disabled={locked || !current.crop || !Object.values(current.crop).some(value => value !== 0)} onclick={reset}>{t('Reset')}</button>
    {:else}
      <p>{t('Crop dimensions are unavailable for an empty image area.')}</p>
    {/if}
  </details>
</section>

<style>
  section { min-width: 0; padding: 12px; border-bottom: 1px solid var(--ok-border); font-size: 12px; }
  summary { cursor: pointer; }
  fieldset { border: 0; margin: 12px 0; padding: 0; min-width: 0; display: grid; gap: 8px; }
  legend { margin-bottom: 8px; color: var(--ok-text-2); }
  label { display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11px; }
  .number { display: flex; align-items: center; gap: 3px; width: 96px; flex: 0 0 96px; }
  input { width: 72px; min-width: 0; padding: 2px 4px; font-size: inherit; }
  .number > span { color: var(--ok-text-2); }
</style>
