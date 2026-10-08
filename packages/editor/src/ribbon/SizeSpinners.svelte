<script lang="ts">
  // The Size group's Height and Width spin boxes (78 × 24 pt) and the Lock
  // Aspect Ratio check box beside them. Shape Format shows only icons before
  // the boxes; Picture Format and Table Layout add "Height:" and "Width:".
  import { cm, emu, getShapeBoundsResolved, isShapeAspectRatioLocked, setShapeAspectRatioLocked, setShapeBounds, type ShapeBounds, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';

  let { labels = false, apply }: {
    labels?: boolean;
    /** Replaces the plain bounds change (tables resize their rows and columns). */
    apply?: (shape: SlideShapeData, bounds: ShapeBounds) => void;
  } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  // The reference desktop app's largest shape dimension, in centimetres.
  const MAX_CM = 5963.92;

  const geometry = $derived.by(() => {
    doc.version;
    return editor.selectedShapes().flatMap((shape) => {
      const bounds = getShapeBoundsResolved(doc.pres, shape);
      return bounds ? [{ shape, bounds }] : [];
    });
  });
  const editable = $derived(geometry.length > 0 && !editor.selectionLocked());
  const common = (field: 'w' | 'h') => {
    const values = geometry.map((item) => item.bounds[field]);
    return values.length > 0 && values.every((value) => value === values[0]) ? Math.round((values[0]! / cm(1)) * 100) / 100 : null;
  };
  const lock = $derived.by(() => {
    doc.version;
    const values = geometry.map((item) => isShapeAspectRatioLocked(item.shape));
    return values.length && values.every((value) => value === values[0]) ? values[0]! : null;
  });
  const lockable = $derived(geometry.length > 0 && geometry.every((item) => item.bounds.w > 0 && item.bounds.h > 0));

  function change(field: 'w' | 'h', input: HTMLInputElement) {
    const restore = () => { input.value = String(common(field) ?? ''); };
    if (!editable || !input.reportValidity() || !Number.isFinite(input.valueAsNumber) || input.valueAsNumber <= 0) { restore(); return; }
    const value = emu(cm(input.valueAsNumber));
    const updates = geometry.map(({ shape, bounds }) => {
      const next = { ...bounds, [field]: value };
      if (isShapeAspectRatioLocked(shape) && lockable) {
        if (field === 'w') next.h = emu((bounds.h * next.w) / bounds.w);
        else next.w = emu((bounds.w * next.h) / bounds.h);
      }
      return { shape, bounds: next };
    });
    if (updates.some((item) => item.bounds.w > cm(MAX_CM) || item.bounds.h > cm(MAX_CM))) { restore(); editor.toast('error', t('The proportional size is too large')); return; }
    doc.transact(t(field === 'w' ? 'Width' : 'Height'), () => {
      for (const item of updates) {
        if (apply) apply(item.shape, item.bounds);
        else setShapeBounds(item.shape, item.bounds);
      }
    });
  }
  function changeLock(input: HTMLInputElement) {
    if (!editable) return;
    doc.setDocumentSetting(() => setShapeAspectRatioLocked(geometry.map((item) => item.shape), input.checked));
  }
</script>

<div class="ctx-spins">
  {#each [['h', 'Height', 'height'], ['w', 'Width', 'width']] as const as [field, name, icon] (field)}
    {@const value = common(field)}
    <label>
      <span class="ctx-spin-label"><Icon name={icon} size={18} />{#if labels}<span>{t(`${name}:`)}</span>{/if}</span>
      <input class="ok-input" type="number" min="0" max={MAX_CM} step="any" aria-label={t(name)} disabled={!editable} value={value ?? ''} placeholder={value === null && geometry.length ? t('Mixed') : undefined} onchange={(event) => change(field, event.currentTarget)} />
    </label>
  {/each}
</div>
<span class="ctx-lock"><input type="checkbox" aria-label={t('Lock Aspect Ratio')} title={t('Lock Aspect Ratio')} checked={lock === true} indeterminate={lock === null && geometry.length > 0} disabled={!editable || !lockable} onchange={(event) => changeLock(event.currentTarget)} /></span>
