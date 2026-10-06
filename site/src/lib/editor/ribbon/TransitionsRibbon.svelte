<script lang="ts">
  import {
    clearSlideTransition,
    getSlides,
    getSlideTransition,
    getSlideTransitionSound,
    setSlideTransition,
    setSlideTransitionSound,
    type SlideData,
    type TransitionEffect,
    type TransitionOptions,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { WRITABLE_TRANSITIONS } from '../core/transition-effects.ts';
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
  const sound = $derived.by(() => {
    doc.version;
    return slides[0] ? getSlideTransitionSound(slides[0]) : null;
  });
  // PowerPoint shows the speed's own duration until one is set.
  const SPEED_MS = { fast: 500, med: 750, slow: 1000 } as const;
  const durationMs = $derived(current?.durationMs ?? SPEED_MS[current?.speed ?? 'med']);
  let soundFile = $state<HTMLInputElement>();
  const timingEditable = $derived(slides.length > 0 && WRITABLE_TRANSITIONS.has(effect));

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
  function changeDuration(ms: number) {
    const { speed: _, ...rest } = current ?? { effect: 'none' };
    apply('Duration', { ...rest, effect: rest.effect as TransitionEffect, durationMs: ms });
  }
  function changeSound(value: string) {
    if (value === 'other') {
      soundFile?.click();
      return;
    }
    doc.transact(t('Sound'), () => {
      for (const slide of slides) setSlideTransitionSound(slide, value === 'stop' ? { kind: 'stop' } : null);
    });
  }
  async function chooseSoundFile(file: File | undefined) {
    if (!file) return;
    const data = new Uint8Array(await file.arrayBuffer());
    try {
      doc.transact(t('Sound'), () => {
        for (const slide of slides) setSlideTransitionSound(slide, { kind: 'play', data, name: file.name });
      });
    } catch (error) {
      editor.toast('error', error instanceof Error ? error.message : String(error));
    }
  }
  // Plays the effect on the editing canvas, as PowerPoint's Preview does.
  function preview() {
    const paint = document.querySelector<HTMLElement>('.canvas-shell .paint');
    if (!paint || effect === 'none') return;
    const dir = current?.direction ?? 'l';
    const from = { l: 'translateX(100%)', r: 'translateX(-100%)', u: 'translateY(100%)', d: 'translateY(-100%)' }[dir[0] as 'l' | 'r' | 'u' | 'd'] ?? 'translateX(100%)';
    const clip = { l: 'inset(0 0 0 100%)', r: 'inset(0 100% 0 0)', u: 'inset(100% 0 0 0)', d: 'inset(0 0 100% 0)' }[dir[0] as 'l' | 'r' | 'u' | 'd'] ?? 'inset(0 0 0 100%)';
    const frames: Keyframe[] =
      effect === 'push' || effect === 'cover' ? [{ transform: from }, { transform: 'none' }]
      : effect === 'wipe' || effect === 'randomBar' ? [{ clipPath: clip }, { clipPath: 'inset(0)' }]
      : effect === 'split' ? [{ clipPath: 'inset(0 50%)' }, { clipPath: 'inset(0)' }]
      : effect === 'cut' ? [{ opacity: 0 }, { opacity: 0, offset: 0.99 }, { opacity: 1 }]
      : [{ opacity: 0 }, { opacity: 1 }];
    paint.animate(frames, { duration: durationMs, easing: 'ease-in-out' });
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

<div class="group">
  <button class="big" disabled={effect === 'none'} onclick={preview}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="16" height="12" rx="1" /><path d="m14 13 7 4-7 4z" fill="currentColor" /></svg>
    <span>{t('Preview')}</span>
  </button>
</div>
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
  <label class="field">
    <svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="6.5" /><path d="M8 4v4l3 2" /></svg>
    {t('Duration:')}
    <input class="seconds" type="number" min="0.01" max="59.99" step="0.25" disabled={!timingEditable || effect === 'none'} value={(durationMs / MS_PER_SECOND).toFixed(2)} onchange={(event) => { const value = event.currentTarget.valueAsNumber; if (Number.isFinite(value) && value > 0) changeDuration(Math.round(value * MS_PER_SECOND)); }} />
  </label>
  <label class="field">
    <svg viewBox="0 0 16 16" aria-hidden="true"><path d="M2 6h3l4-3v10l-4-3H2zM11 5.5a3.5 3.5 0 0 1 0 5M12.5 3.5a6 6 0 0 1 0 9" /></svg>
    {t('Sound:')}
    <select disabled={slides.length === 0} value={sound?.kind === 'play' ? 'play' : sound?.kind ?? 'none'} onchange={(event) => changeSound(event.currentTarget.value)}>
      <option value="none">{t('[No Sound]')}</option>
      <option value="stop">{t('[Stop Previous Sound]')}</option>
      {#if sound?.kind === 'play'}<option value="play">{sound.name}</option>{/if}
      <option value="other">{t('Other Sound…')}</option>
    </select>
  </label>
  <input bind:this={soundFile} class="file" type="file" accept=".wav,audio/wav" tabindex="-1" aria-label={t('Other Sound…')} onchange={(event) => { void chooseSoundFile(event.currentTarget.files?.[0]); event.currentTarget.value = ''; }} />
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
  .field svg { width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.1; }
  select { min-width: 140px; font: inherit; padding: 1px 4px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-panel); color: var(--ok-text); }
  .file { position: absolute; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
  .seconds { width: 64px; font: inherit; padding: 1px 4px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-panel); color: var(--ok-text); }
</style>
