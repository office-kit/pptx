<script lang="ts">
  import {
    getSlideAnimations,
    removeSlideAnimation,
    setShapeAnimation,
    updateSlideAnimation,
    type AnimationEffect,
    type AnimationStartCondition,
  } from '@office-kit/pptx';
  import { tick } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  // Mac PowerPoint's animation gallery, limited to the presets the library
  // writes. The star color marks the class, as it does there: green entrance,
  // gold emphasis, red exit.
  const GALLERY: ReadonlyArray<{ effect: AnimationEffect; label: string; kind: 'entrance' | 'emphasis' | 'exit' }> = [
    { effect: 'appear', label: 'Appear', kind: 'entrance' },
    { effect: 'fadeIn', label: 'Fade', kind: 'entrance' },
    { effect: 'flyIn', label: 'Fly In', kind: 'entrance' },
    { effect: 'zoomIn', label: 'Zoom', kind: 'entrance' },
    { effect: 'spin', label: 'Spin', kind: 'emphasis' },
    { effect: 'disappear', label: 'Disappear', kind: 'exit' },
    { effect: 'fadeOut', label: 'Fade Out', kind: 'exit' },
    { effect: 'flyOut', label: 'Fly Out', kind: 'exit' },
    { effect: 'zoomOut', label: 'Zoom Out', kind: 'exit' },
  ];
  const STARTS: ReadonlyArray<readonly [AnimationStartCondition, string]> = [
    ['click', 'On Click'],
    ['withPrevious', 'With Previous'],
    ['afterPrevious', 'After Previous'],
  ];
  const MS_PER_SECOND = 1000;

  const editor = getEditor();
  const doc = editor.doc;
  const shapeId = $derived(selectedShapeId(doc.selection));
  // The first effect on the selected shape is the one the gallery shows and
  // replaces, like PowerPoint's single-effect gallery.
  const step = $derived.by(() => {
    doc.version;
    const slide = doc.currentSlide;
    if (!slide || shapeId === null) return null;
    return getSlideAnimations(slide).find((item) => item.id !== null && item.targetShapeIds.includes(shapeId)) ?? null;
  });

  function choose(effect: AnimationEffect | null) {
    const slide = doc.currentSlide;
    const shape = editor.selectedShapes()[0];
    if (!slide || !shape) return;
    doc.transact(t('Animation'), () => {
      if (effect === null) {
        for (const item of getSlideAnimations(slide)) {
          if (item.id !== null && shapeId !== null && item.targetShapeIds.includes(shapeId)) removeSlideAnimation(slide, item.id);
        }
      } else if (step?.id != null) updateSlideAnimation(slide, step.id, { effect });
      else setShapeAnimation(shape, { effect });
    });
  }
  function patch(value: { start?: AnimationStartCondition; durationMs?: number }) {
    const slide = doc.currentSlide;
    if (!slide || step?.id == null) return;
    const id = step.id;
    doc.transact(t('Animation'), () => updateSlideAnimation(slide, id, value));
  }
  async function showPane() {
    // It sits with the shape's Size & Properties sections in the Format pane.
    editor.showShapeFormat('size');
    await tick();
    const pane = document.querySelector<HTMLElement>('section[data-animation-pane]');
    pane?.scrollIntoView({ block: 'nearest' });
    pane?.focus();
  }
</script>

<div class="group gallery" role="radiogroup" aria-label={t('Animation')}>
  <button role="radio" aria-checked={shapeId !== null && step === null} disabled={shapeId === null} onclick={() => choose(null)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.4 6.7 19.4l1.2-6L3.4 9.3l6-.7z" /><path d="M4 20 20 4" /></svg>
    <span>{t('None')}</span>
  </button>
  {#each GALLERY as item (item.effect)}
    <button role="radio" class={item.kind} aria-checked={step?.effect === item.effect} disabled={shapeId === null} onclick={() => choose(item.effect)}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path class="star" d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.4 6.7 19.4l1.2-6L3.4 9.3l6-.7z" /></svg>
      <span>{t(item.label)}</span>
    </button>
  {/each}
</div>
<div class="group">
  <button class="big" disabled={step?.id == null} onclick={() => editor.runOrPrompt('setShapeAnimation')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 3l2 4.4 4.7.5-3.5 3.2.9 4.6L10 13.4l-4.1 2.3.9-4.6-3.5-3.2 4.7-.5z" /><circle cx="17" cy="17" r="3" /></svg>
    <span>{t('Effect Options')}</span>
  </button>
  <button class="big" onclick={showPane}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.4 6.7 19.4l1.2-6L3.4 9.3l6-.7z" /><path d="M16 15h6M16 18h6M16 21h6" /></svg>
    <span>{t('Animation Pane')}</span>
  </button>
</div>
<div class="group timing">
  <label>{t('Start:')}
    <select aria-label={t('Start')} disabled={step?.id == null} value={step?.start ?? ''} onchange={(event) => patch({ start: event.currentTarget.value as AnimationStartCondition })}>
      {#if step?.id == null}<option value=""></option>{/if}
      {#each STARTS as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
    </select>
  </label>
  <label>{t('Duration:')}
    <input type="number" min="0.01" step="0.25" aria-label={t('Duration')} disabled={step?.id == null} value={step?.durationMs != null ? (step.durationMs / MS_PER_SECOND).toFixed(2) : ''} onchange={(event) => { const value = event.currentTarget.valueAsNumber; if (Number.isFinite(value) && value > 0) patch({ durationMs: Math.round(value * MS_PER_SECOND) }); }} />
  </label>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  button { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; padding: 3px 4px; font: inherit; font-size: 11px; line-height: 1.15; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:not(:disabled):hover, button[aria-checked='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  button:disabled, label:has(:disabled) { opacity: 0.45; cursor: default; }
  .gallery button { width: 58px; min-height: 60px; }
  svg { width: 32px; height: 32px; flex-shrink: 0; fill: none; stroke: currentColor; stroke-width: 1.1; }
  .entrance .star { fill: #3f9b4a; stroke: #2d7a37; }
  .emphasis .star { fill: #d9a520; stroke: #a87c10; }
  .exit .star { fill: #c8423b; stroke: #9c2d27; }
  .big { min-width: 52px; max-width: 72px; min-height: 66px; text-align: center; }
  .timing { flex-direction: column; align-items: flex-end; justify-content: center; gap: 8px; font-size: 12px; }
  label { display: flex; align-items: center; gap: 6px; white-space: nowrap; }
  select, input { width: 120px; font: inherit; padding: 1px 4px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-panel); color: var(--ok-text); }
</style>
