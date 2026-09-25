<script lang="ts">
  import { untrack } from 'svelte';
  import { cm, emu, getSlideSize, getShapeBoundsResolved, getShapeId, getShapeRotation, getShapeFlip, setShapeBounds, type ShapeBounds } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const maxDimension = 5963.92; // Mac PowerPoint's Size and Position input limit, in cm.
  const locked = $derived(editor.selectionLocked());
  const geometry = $derived.by(() => {
    doc.version;
    const items = [];
    for (const shape of editor.selectedShapes()) {
      const bounds = getShapeBoundsResolved(doc.pres, shape);
      if (!bounds) return [];
      items.push({ shape, bounds });
    }
    return items;
  });
  let originals = $state.raw(new Map<number, ShapeBounds>());
  $effect(() => {
    doc.selection;
    doc.pres;
    // Scaling percentages use the dimensions when the objects were selected.
    originals = untrack(() => new Map(geometry.map(item => [getShapeId(item.shape), { ...item.bounds }])));
  });
  let sizeOpen = $state(true);
  let positionOpen = $state(true);
  let lockAspectRatio = $state(false);
  let origins = $state({ x: 'corner', y: 'corner' });
  const slideSize = $derived.by(() => { doc.version; return getSlideSize(doc.pres); });
  function originOffset(field: keyof ShapeBounds) {
    if (field === 'x' && origins.x === 'center') return (slideSize?.width ?? 0) / 2;
    if (field === 'y' && origins.y === 'center') return (slideSize?.height ?? 0) / 2;
    return 0;
  }
  function changeOrigin(axis: 'x' | 'y', input: HTMLSelectElement) {
    if (editor.selectionLocked()) { input.value = origins[axis]; return; }
    const next = input.value;
    const half = (axis === 'x' ? slideSize?.width : slideSize?.height) ?? 0;
    const delta = (next === 'center' ? half / 2 : 0) - originOffset(axis);
    if (geometry.some(item => Math.abs(item.bounds[axis] + delta) > cm(maxDimension))) {
      input.value = origins[axis]; return;
    }
    // Mac PowerPoint keeps the entered number and moves the object when the
    // origin changes. Undo restores geometry but keeps this panel preference.
    origins[axis] = next;
    doc.transact(t('Set bounds'), () => {
      for (const item of geometry) setShapeBounds(item.shape, { ...item.bounds, [axis]: emu(item.bounds[axis] + delta) });
    });
  }
  let rotationInput = $state<HTMLInputElement>();
  $effect(() => {
    if (editor.rotationFocusRequested && !sizeOpen) { sizeOpen = true; return; }
    if (editor.rotationFocusRequested && rotationInput) {
      rotationInput.scrollIntoView({ block: 'nearest' });
      rotationInput.focus(); rotationInput.select();
      editor.rotationFocusRequested = false;
    }
  });
  const canLockAspectRatio = $derived(geometry.length > 0 && geometry.every(item => item.bounds.w > 0 && item.bounds.h > 0));
  const round = (value: number) => Math.round(value * 100) / 100;
  function common(values: number[]): number | null {
    return values.length > 0 && values.every(value => value === values[0]) ? values[0]! : null;
  }
  const bounds = $derived.by(() => {
    const value = (field: keyof ShapeBounds) => {
      const result = common(geometry.map(item => item.bounds[field]));
      return result === null ? null : round((result - originOffset(field)) / cm(1));
    };
    return { x: value('x'), y: value('y'), w: value('w'), h: value('h') };
  });
  const rotation = $derived.by(() => { doc.version; return common(editor.selectedShapes().map(getShapeRotation)); });
  const scales = $derived.by(() => {
    const value = (field: 'w' | 'h') => common(geometry.map(item => {
      const original = originals.get(getShapeId(item.shape))?.[field];
      return original ? round(item.bounds[field] / original * 100) : 100;
    }));
    return { w: value('w'), h: value('h') };
  });
  const scaleLimits = $derived.by(() => {
    const limit = (field: 'w' | 'h') => {
      let maximum = Infinity;
      for (const bounds of originals.values()) {
        if (bounds[field] > 0) maximum = Math.min(maximum, cm(maxDimension) / bounds[field] * 100);
      }
      return maximum;
    };
    return { w: limit('w'), h: limit('h') };
  });
  const flips = $derived.by(() => {
    doc.version;
    const values = editor.selectedShapes().map(getShapeFlip);
    const value = (axis: 'horizontal' | 'vertical') => {
      const states = new Set(values.map(item => item?.[axis] ?? false));
      return states.size > 1 ? null : states.has(true);
    };
    return { horizontal: value('horizontal'), vertical: value('vertical') };
  });

  function change(field: keyof ShapeBounds, input: HTMLInputElement, scale = false) {
    const restore = () => { const value = scale ? scales[field as 'w' | 'h'] : bounds[field]; input.value = value === null ? '' : String(value); };
    if (!geometry.length || editor.selectionLocked() || !input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { restore(); return; }
    const value = input.valueAsNumber;
    const updates = geometry.map(item => {
      const next = { ...item.bounds, [field]: scale ? emu((originals.get(getShapeId(item.shape))?.[field] ?? item.bounds[field]) * value / 100) : emu(cm(value) + originOffset(field)) };
      if (lockAspectRatio && canLockAspectRatio) {
        if (field === 'w') next.h = emu(item.bounds.h * next.w / item.bounds.w);
        if (field === 'h') next.w = emu(item.bounds.w * next.h / item.bounds.h);
      }
      return { shape: item.shape, bounds: next };
    });
    if (updates.some(item => item.bounds.w > cm(maxDimension) || item.bounds.h > cm(maxDimension))) {
      restore(); editor.toast('error', t('The proportional size is too large')); return;
    }
    doc.transact(t('Set bounds'), () => { for (const item of updates) setShapeBounds(item.shape, item.bounds); });
  }
  function rotate(input: HTMLInputElement) {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { input.value = rotation === null ? '' : String(rotation); return; }
    editor.invoke('setShapeRotation', { degrees: input.valueAsNumber });
  }
</script>

{#if geometry.length}
  <div class="geometry">
    <details bind:open={sizeOpen}>
      <summary>{t('Size')}</summary>
      <div class="fields">
        {#each [['h', 'Height'], ['w', 'Width']] as [field, label]}
          {@const axis = field as 'h' | 'w'}
          <label><span>{t(label!)}</span><span class="number"><input class="ok-input" type="number" aria-label={t(label!)} disabled={locked} min="0" max={maxDimension} step="any" value={bounds[axis] ?? ''} placeholder={bounds[axis] === null ? t('Mixed') : undefined} onchange={event => change(axis, event.currentTarget)} /><span>cm</span></span></label>
        {/each}
        <label><span>{t('Rotation')}</span><span class="number"><input bind:this={rotationInput} class="ok-input" type="number" aria-label={t('Rotation')} disabled={locked} min="-3600" max="3600" step="any" value={rotation ?? ''} placeholder={rotation === null ? t('Mixed') : undefined} onchange={event => rotate(event.currentTarget)} /><span>°</span></span></label>
        {#each [['h', 'Scale Height'], ['w', 'Scale Width']] as [field, label]}
          {@const axis = field as 'h' | 'w'}
          <label><span>{t(label!)}</span><span class="number"><input class="ok-input" type="number" aria-label={t(label!)} disabled={locked || geometry.some(item => !originals.get(getShapeId(item.shape))?.[axis])} min="1" max={Number.isFinite(scaleLimits[axis]) ? scaleLimits[axis] : undefined} step="any" value={scales[axis] ?? ''} placeholder={scales[axis] === null ? t('Mixed') : undefined} onchange={event => change(axis, event.currentTarget, true)} /><span>%</span></span></label>
        {/each}
        <label class="check"><input type="checkbox" bind:checked={lockAspectRatio} disabled={locked || !canLockAspectRatio} /><span>{t('Lock aspect ratio')}</span></label>
      </div>
    </details>
    <details bind:open={positionOpen}>
      <summary>{t('Position')}</summary>
      <div class="fields">
        {#each [['x', 'Horizontal position'], ['y', 'Vertical position']] as [field, label]}
          {@const axis = field as 'x' | 'y'}
          <label><span>{t(label!)}</span><span class="number"><input class="ok-input" type="number" aria-label={t(label!)} disabled={locked} min={-maxDimension - originOffset(axis) / cm(1)} max={maxDimension - originOffset(axis) / cm(1)} step="any" value={bounds[axis] ?? ''} placeholder={bounds[axis] === null ? t('Mixed') : undefined} onchange={event => change(axis, event.currentTarget)} /><span>cm</span></span></label>
          <label><span>{t('From')}</span><select class="ok-input" aria-label={t(axis === 'x' ? 'Horizontal position from' : 'Vertical position from')} disabled={locked} value={origins[axis]} onchange={event => changeOrigin(axis, event.currentTarget)}><option value="corner">{t('Top Left Corner')}</option><option value="center">{t('Center')}</option></select></label>
        {/each}
      </div>
    </details>
    <div class="fields flips">
      {#each ['horizontal', 'vertical'] as axis}
        {@const value = axis === 'horizontal' ? flips.horizontal : flips.vertical}
        <label class="check"><input type="checkbox" disabled={locked} checked={value ?? false} indeterminate={value === null} onchange={event => editor.invoke('setShapeFlip', { options: { [axis]: event.currentTarget.checked } })} /><span>{t(axis === 'horizontal' ? 'Flip horizontally' : 'Flip vertically')}{value === null ? ` (${t('Mixed')})` : ''}</span></label>
      {/each}
    </div>
  </div>
{/if}

<style>
  .geometry { font-size: 11px; margin: 0 -10px; }
  summary { padding: 4px 8px; background: var(--ok-hover); cursor: pointer; }
  .fields { display: flex; flex-direction: column; gap: 8px; padding: 12px; }
  label { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 20px; }
  .number { display: flex; align-items: center; gap: 3px; width: 96px; flex: 0 0 96px; }
  .number .ok-input { width: 72px; min-width: 0; padding: 2px 4px; font-size: inherit; }
  .number > span { color: var(--ok-text-2); }
  .check { justify-content: flex-start; }
  .check input { margin: 0; }
  .flips { padding-top: 6px; }
</style>
