<script lang="ts">
  import {
    clearSlideTransition,
    getSlides,
    getSlideTransition,
    setSlideTransition,
    type SlideData,
    type TransitionEffect,
    type TransitionOptions,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  // Mac PowerPoint's Transitions gallery, limited to the effects the library
  // writes (Morph, Reveal and Shape are PowerPoint 2010+ extensions).
  const GALLERY: ReadonlyArray<readonly [TransitionEffect, string]> = [
    ['none', 'None'],
    ['fade', 'Fade'],
    ['push', 'Push'],
    ['wipe', 'Wipe'],
    ['split', 'Split'],
    ['cut', 'Cut'],
    ['randomBar', 'Random Bars'],
    ['cover', 'Cover'],
    ['pull', 'Uncover'],
  ];
  // Every token setSlideTransition writes; a deck can carry others (p14
  // extensions), which the timing controls must not rewrite.
  const WRITABLE: ReadonlySet<string> = new Set<TransitionEffect>(['none', 'blinds', 'checker', 'circle', 'dissolve', 'comb', 'cover', 'cut', 'diamond', 'fade', 'newsflash', 'plus', 'pull', 'push', 'random', 'randomBar', 'split', 'strips', 'wedge', 'wheel', 'wipe', 'zoom']);
  const MS_PER_SECOND = 1000;

  const editor = getEditor();
  const doc = editor.doc;
  const slides = $derived.by(() => {
    doc.version;
    const all = getSlides(doc.pres);
    return selectedSlideIndices(doc.selection).flatMap((index) => (all[index] ? [all[index]!] : []));
  });
  const current = $derived.by(() => {
    doc.version;
    return slides[0] ? getSlideTransition(slides[0]) : null;
  });
  const effect = $derived(current?.effect ?? 'none');
  const onClick = $derived(current?.advanceOnClick ?? true);
  const afterMs = $derived(current?.advanceAfterMs);
  const timingEditable = $derived(slides.length > 0 && WRITABLE.has(effect));

  function apply(label: string, value: TransitionOptions, targets: readonly SlideData[] = slides) {
    doc.transact(t(label), () => {
      for (const slide of targets) {
        if (value.effect === 'none' && value.advanceOnClick !== false && value.advanceAfterMs === undefined) clearSlideTransition(slide);
        else setSlideTransition(slide, value);
      }
    });
  }
  // A new effect keeps the slide's timing but not the old effect's direction
  // or speed, which belong to that effect.
  function chooseEffect(token: TransitionEffect) {
    apply('Slide transition', { effect: token, advanceOnClick: onClick, ...(afterMs !== undefined ? { advanceAfterMs: afterMs } : {}) });
  }
  function changeTiming(advanceOnClick: boolean, advanceAfterMs: number | undefined, targets: readonly SlideData[] = slides) {
    const { advanceAfterMs: _, ...rest } = current ?? { effect: 'none' };
    apply(targets === slides ? 'Slide transition' : 'Apply To All', {
      ...rest,
      effect: rest.effect as TransitionEffect,
      advanceOnClick,
      ...(advanceAfterMs !== undefined ? { advanceAfterMs } : {}),
    }, targets);
  }
</script>

<div class="group gallery" role="radiogroup" aria-label={t('Transition to This Slide')}>
  {#each GALLERY as [token, label] (token)}
    <button role="radio" aria-checked={effect === token} disabled={slides.length === 0} onclick={() => chooseEffect(token)}>
      <svg viewBox="0 0 40 26" aria-hidden="true">
        <rect x="1" y="1" width="38" height="24" class="slide" />
        {#if token === 'fade'}<rect x="8" y="5" width="28" height="18" class="next half" />
        {:else if token === 'push'}<rect x="1" y="13" width="38" height="12" class="next" /><path d="M20 21v-6m-3 3 3-3 3 3" />
        {:else if token === 'wipe'}<rect x="20" y="1" width="19" height="24" class="next" /><path d="M28 13h-6m3-3-3 3 3 3" />
        {:else if token === 'split'}<rect x="14" y="1" width="12" height="24" class="next" /><path d="M10 13H4m3-3-3 3 3 3M30 13h6m-3-3 3 3-3 3" />
        {:else if token === 'cut'}<rect x="10" y="6" width="28" height="18" class="next" />
        {:else if token === 'randomBar'}<path d="M6 1v24M10 1v24M13 1v24M18 1v24M22 1v24M27 1v24M31 1v24M35 1v24" />
        {:else if token === 'cover'}<rect x="12" y="1" width="27" height="24" class="next" />
        {:else if token === 'pull'}<rect x="1" y="1" width="26" height="24" class="next half" />{/if}
      </svg>
      <span>{t(label)}</span>
    </button>
  {/each}
</div>
<div class="group">
  <button class="big" disabled={slides.length === 0} onclick={() => editor.runOrPrompt('setSlideTransition')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="1" /><path d="m15 10 2 2-2 2" /></svg>
    <span>{t('Effect Options')}</span>
  </button>
</div>
<div class="group timing">
  <label><input type="checkbox" checked={onClick} disabled={!timingEditable} onchange={(event) => changeTiming(event.currentTarget.checked, afterMs)} />{t('On Mouse Click')}</label>
  <div class="after">
    <label><input type="checkbox" checked={afterMs !== undefined} disabled={!timingEditable} onchange={(event) => changeTiming(onClick, event.currentTarget.checked ? 0 : undefined)} />{t('After:')}</label>
    <input class="seconds" type="number" min="0" step="0.01" aria-label={t('Advance after (seconds)')} disabled={!timingEditable || afterMs === undefined} value={((afterMs ?? 0) / MS_PER_SECOND).toFixed(2)} onchange={(event) => { const value = event.currentTarget.valueAsNumber; if (Number.isFinite(value) && value >= 0) changeTiming(onClick, Math.round(value * MS_PER_SECOND)); }} />
  </div>
</div>
<div class="group">
  <button class="big" disabled={!timingEditable} onclick={() => changeTiming(onClick, afterMs, getSlides(doc.pres))}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="7" y="3" width="14" height="11" rx="1" /><rect x="3" y="8" width="14" height="11" rx="1" /><path d="m14 17 3 2-3 2" /></svg>
    <span>{t('Apply To All')}</span>
  </button>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  button { display: flex; flex-direction: column; align-items: center; gap: 2px; padding: 3px 4px; font: inherit; font-size: 11px; line-height: 1.15; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:not(:disabled):hover, button[aria-checked='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  button:disabled, label:has(input:disabled) { opacity: 0.45; cursor: default; }
  .gallery button { width: 64px; min-height: 60px; }
  .gallery svg { width: 48px; height: 32px; fill: none; stroke: currentColor; stroke-width: 1; }
  .gallery .slide { fill: var(--ok-panel); }
  .gallery .next { fill: color-mix(in srgb, var(--ok-accent) 35%, transparent); stroke: none; }
  .gallery .next.half { fill: color-mix(in srgb, var(--ok-accent) 18%, transparent); }
  .big { min-width: 52px; max-width: 72px; min-height: 66px; text-align: center; }
  .big svg { width: 32px; height: 32px; flex-shrink: 0; fill: none; stroke: currentColor; stroke-width: 1.1; }
  .timing { flex-direction: column; align-items: flex-start; justify-content: center; gap: 8px; font-size: 12px; }
  .after { display: flex; align-items: center; gap: 6px; }
  label { display: flex; align-items: center; gap: 6px; white-space: nowrap; }
  input[type='checkbox'] { margin: 0; accent-color: var(--ok-accent); }
  .seconds { width: 64px; font: inherit; padding: 1px 4px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-panel); color: var(--ok-text); }
</style>
