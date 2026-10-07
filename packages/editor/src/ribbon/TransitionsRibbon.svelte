<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // Mac PowerPoint 16's Transitions tab: Preview | the Transition Styles
  // gallery with Effect Options beside it | Duration, Sound, On Mouse Click,
  // After and Apply To All as one timing group. Measured at 1512 and 1200 pt
  // windows: the gallery shows as many 92 pt tiles as fit (10 and 6), and
  // nothing else changes size.
  import {
    clearSlideTransition,
    getSlides,
    getSlideTransition,
    getSlideTransitionSound,
    setSlideTransition,
    setSlideTransitionSound,
    type SlideData,
    type TransitionOptions,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { writableTransition } from '../core/transition-effects.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import RibbonGallery from './RibbonGallery.svelte';
  import { placeBelowTrigger } from './place-menu.ts';
  import {
    optionMatches,
    shownDurationMs,
    tileOfTransition,
    TRANSITION_OPTIONS,
    TRANSITION_TILES,
    type TransitionChoice,
    type TransitionTile,
  } from './transition-gallery.ts';

  const MS_PER_SECOND = 1000;
  // Everything but the gallery's tiles takes 576 px of the command row at
  // PowerPoint's sizes (Preview, the gallery's arrow columns, Effect Options,
  // the timing group, the rules between them and a 16 px end margin).
  const FIXED_WIDTH = 576;
  const END_MARGIN = 16;
  const TILE_WIDTH = 92;
  const MIN_TILES = 3;

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
  const applied = $derived(current?.effect ?? 'none');
  const onClick = $derived(current?.advanceOnClick ?? true);
  const afterMs = $derived(current?.advanceAfterMs);
  const sound = $derived.by(() => {
    doc.version;
    return slides[0] ? getSlideTransitionSound(slides[0]) : null;
  });
  const durationMs = $derived(shownDurationMs(current));
  let soundFile = $state<HTMLInputElement>();
  // What a timing edit writes back: the slide's transition as it stands, or
  // null when its effect is one the library cannot write.
  const writable = $derived(writableTransition(current));
  const timingEditable = $derived(slides.length > 0 && writable !== null);
  const checkedTile = $derived(tileOfTransition(current));
  const options = $derived(checkedTile ? (TRANSITION_OPTIONS[checkedTile] ?? []) : []);
  let width = $state(typeof window === 'undefined' ? Number.POSITIVE_INFINITY : window.innerWidth);
  // Longer labels (Japanese) widen the timing group, so the row measures what
  // it needs besides the tiles rather than trusting the English widths.
  let fixed = $state(FIXED_WIDTH);
  let row = $state<HTMLDivElement>();
  const tiles = $derived(Math.max(MIN_TILES, Math.floor((width - fixed) / TILE_WIDTH)));
  $effect(() => {
    getLocale();
    const shown = tiles;
    const last = row?.lastElementChild;
    if (!row || !last) return;
    const used = last.getBoundingClientRect().right - row.getBoundingClientRect().left;
    fixed = Math.max(FIXED_WIDTH, Math.ceil(used) - shown * TILE_WIDTH + END_MARGIN);
  });
  let menuOpen = $state(false);
  const label = (item: { en: string; ja: string }) => (getLocale() === 'ja' ? item.ja : item.en);

  function apply(name: string, value: TransitionOptions, targets: readonly SlideData[] = slides) {
    doc.transact(t(name), () => {
      for (const slide of targets) {
        if (value.effect === 'none' && value.advanceOnClick !== false && value.advanceAfterMs === undefined && value.durationMs === undefined) clearSlideTransition(slide);
        else setSlideTransition(slide, value);
      }
    });
  }
  const timing = () => ({ advanceOnClick: onClick, ...(afterMs !== undefined ? { advanceAfterMs: afterMs } : {}) });
  // A new effect keeps the slide's timing but not the old effect's direction
  // or duration, which belong to that effect: it takes the effect's own
  // default duration, as PowerPoint does.
  function chooseTile(item: TransitionTile) {
    const duration = item.choice.effect === 'none' ? {} : { durationMs: item.durationMs };
    apply('Slide transition', { ...item.choice, ...duration, ...timing() });
  }
  // An Effect Options item keeps the effect's duration as well, including
  // one written as a speed alone.
  function chooseOption(choice: TransitionChoice) {
    menuOpen = false;
    apply('Effect Options', { ...choice, ...timing(), durationMs });
  }
  function changeDuration(ms: number) {
    if (!writable) return;
    const { speed: _, ...rest } = writable;
    apply('Duration', { ...rest, durationMs: ms });
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
  // Plays the effect on the editing canvas, as PowerPoint's Preview does. The
  // 3-D and particle effects of PowerPoint 2010+ are approximated by the
  // nearest flat motion: a slide-in, a wipe, a split, a zoom or a fade.
  const SLIDES = new Set(['push', 'cover', 'pan', 'gallery', 'conveyor', 'ferris', 'vortex']);
  const WIPES = new Set(['wipe', 'randomBar', 'strips', 'blinds', 'checker', 'comb', 'reveal', 'glitter']);
  const SPLITS = new Set(['split', 'doors', 'window', 'shred']);
  const ZOOMS = new Set(['zoom', 'warp', 'flythrough', 'newsflash', 'circle', 'diamond', 'plus', 'prism', 'ripple']);
  function preview() {
    const paint = editor.shell?.querySelector<HTMLElement>('.canvas-shell .paint');
    if (!paint || applied === 'none') return;
    const dir = current?.direction ?? 'l';
    const side = (['l', 'r', 'u', 'd'] as const).find((token) => dir.startsWith(token)) ?? 'l';
    const from = { l: 'translateX(100%)', r: 'translateX(-100%)', u: 'translateY(100%)', d: 'translateY(-100%)' }[side];
    const clip = { l: 'inset(0 0 0 100%)', r: 'inset(0 100% 0 0)', u: 'inset(100% 0 0 0)', d: 'inset(0 0 100% 0)' }[side];
    const frames: Keyframe[] =
      SLIDES.has(applied) ? [{ transform: from }, { transform: 'none' }]
      : WIPES.has(applied) ? [{ clipPath: clip }, { clipPath: 'inset(0)' }]
      : SPLITS.has(applied) ? [{ clipPath: dir === 'horz' || current?.orientation === 'horz' ? 'inset(50% 0)' : 'inset(0 50%)' }, { clipPath: 'inset(0)' }]
      : ZOOMS.has(applied) ? [{ transform: dir === 'out' ? 'scale(1.6)' : 'scale(0.2)', opacity: 0 }, { transform: 'none', opacity: 1 }]
      : applied === 'cut' ? [{ opacity: 0 }, { opacity: 0, offset: 0.99 }, { opacity: 1 }]
      : [{ opacity: 0 }, { opacity: 1 }];
    paint.animate(frames, { duration: durationMs, easing: 'ease-in-out' });
  }
  function changeTiming(advanceOnClick: boolean, advanceAfterMs: number | undefined, targets: readonly SlideData[] = slides) {
    if (!writable) return;
    const { advanceAfterMs: _, ...rest } = writable;
    apply(targets === slides ? 'Slide transition' : 'Apply To All', {
      ...rest,
      advanceOnClick,
      ...(advanceAfterMs !== undefined ? { advanceAfterMs } : {}),
    }, targets);
  }
</script>

<svelte:window onkeydown={(event) => { if (event.key === 'Escape') menuOpen = false; }} onpointerdown={(event) => { if (menuOpen && !(eventTarget(event) as Element).closest?.('.anchor')) menuOpen = false; }} />

{#snippet art(token: string)}
  <svg class="art" viewBox="0 0 40 26" aria-hidden="true">
    <rect x="1" y="1" width="38" height="24" class="slide" />
    {#if token === 'Fade'}<rect x="8" y="5" width="28" height="18" class="next half" />
    {:else if token === 'Push'}<rect x="1" y="13" width="38" height="12" class="next" /><path d="M20 21v-6m-3 3 3-3 3 3" />
    {:else if token === 'Wipe'}<rect x="20" y="1" width="19" height="24" class="next" /><path d="M28 13h-6m3-3-3 3 3 3" />
    {:else if token === 'Split'}<rect x="14" y="1" width="12" height="24" class="next" /><path d="M10 13H4m3-3-3 3 3 3M30 13h6m-3-3 3 3-3 3" />
    {:else if token === 'Cut'}<rect x="10" y="6" width="28" height="18" class="next" />
    {:else if token === 'Random Bars'}<path d="M6 1v24M10 1v24M13 1v24M18 1v24M22 1v24M27 1v24M31 1v24M35 1v24" />
    {:else if token === 'Shape'}<circle cx="20" cy="13" r="8" class="next" />
    {:else if token === 'Cover'}<rect x="12" y="1" width="27" height="24" class="next" />
    {:else if token === 'Uncover'}<rect x="1" y="1" width="26" height="24" class="next half" />
    {:else if token === 'Checkerboard'}<path d="M1 1h8v8H1zM17 1h8v8h-8zM9 9h8v8H9zM25 9h8v8h-8zM1 17h8v8H1zM17 17h8v8h-8z" class="next" />
    {:else if token === 'Blinds'}<path d="M1 4h38M1 10h38M1 16h38M1 22h38" />
    {:else if token === 'Clock'}<path d="M20 13V2a11 11 0 0 1 11 11z" class="next" />
    {:else if token === 'Comb'}<path d="M1 1h38v6H1zM1 13h38v6H1z" class="next" />
    {:else if token === 'Zoom'}<rect x="12" y="8" width="16" height="10" class="next" />
    {:else if token !== 'None'}<rect x="10" y="6" width="22" height="14" class="next half" />{/if}
  </svg>
{/snippet}

<div class="transitions" lang={getLocale()} bind:this={row} bind:clientWidth={width}>
  <section class="cluster" role="group" aria-label={t('Preview')}>
    <button class="big" style:--w="46px" disabled={applied === 'none'} onclick={preview}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="16" height="12" rx="1" /><path d="m14 13 7 4-7 4z" fill="currentColor" /></svg>
      <span>{t('Preview')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Transition to This Slide')}>
    <RibbonGallery label={t('Transition Styles')} items={TRANSITION_TILES} checked={checkedTile} visible={tiles} tileWidth={TILE_WIDTH} disabled={slides.length === 0} name={label} choose={chooseTile}>
      {#snippet tile(item)}{@render art(item.key)}{/snippet}
    </RibbonGallery>
    <div class="anchor">
      <button class="big" style:--w="50px" aria-haspopup="menu" aria-expanded={menuOpen} disabled={slides.length === 0 || options.length === 0} onclick={() => (menuOpen = !menuOpen)}>
        <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="1" /><path d="M3 12h18" /><path d="M12 18v-4m-2 2 2-2 2 2" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
        <span>{t('Effect Options')}</span>
      </button>
      {#if menuOpen}
        <div class="menu" role="menu" aria-label={t('Effect Options')} use:placeBelowTrigger>
          {#each options as item (item.en)}
            <button role="menuitemradio" aria-checked={current !== null && optionMatches(item.choice, current)} onclick={() => chooseOption(item.choice)}>
              <span class="check" aria-hidden="true">✓</span>
              <svg class="thumb" viewBox="0 0 40 26" aria-hidden="true"><rect x="1" y="1" width="38" height="24" class="slide" /><rect x="1" y="1" width="38" height="24" class="next half" />{#if item.arrow !== undefined}<path d="M14 13h12m-4-4 4 4-4 4" transform="rotate({-item.arrow} 20 13)" />{/if}</svg>
              <span>{label(item)}</span>
            </button>
          {/each}
        </div>
      {/if}
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Timing')}>
    <div class="timing">
      <label class="field">
        <svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="6.5" /><path d="M8 4v4l3 2" /></svg>
        <span>{t('Duration:')}</span>
        <input class="seconds" type="number" min="0.01" max="59.99" step="0.25" disabled={!timingEditable} value={(durationMs / MS_PER_SECOND).toFixed(2)} onchange={(event) => { const value = event.currentTarget.valueAsNumber; if (Number.isFinite(value) && value > 0) changeDuration(Math.round(value * MS_PER_SECOND)); }} />
      </label>
      <label class="check-row"><input type="checkbox" checked={onClick} disabled={!timingEditable} onchange={(event) => changeTiming(event.currentTarget.checked, afterMs)} />{t('On Mouse Click')}</label>
      <label class="field">
        <svg viewBox="0 0 16 16" aria-hidden="true"><path d="M2 6h3l4-3v10l-4-3H2zM11 5.5a3.5 3.5 0 0 1 0 5M12.5 3.5a6 6 0 0 1 0 9" /></svg>
        <span>{t('Sound:')}</span>
        <select disabled={slides.length === 0} value={sound?.kind === 'play' ? 'play' : sound?.kind ?? 'none'} onchange={(event) => changeSound(event.currentTarget.value)}>
          <option value="none">{t('[No Sound]')}</option>
          <option value="stop">{t('[Stop Previous Sound]')}</option>
          {#if sound?.kind === 'play'}<option value="play">{sound.name}</option>{/if}
          <option value="other">{t('Other Sound…')}</option>
        </select>
      </label>
      <div class="check-row">
        <label><input type="checkbox" checked={afterMs !== undefined} disabled={!timingEditable} onchange={(event) => changeTiming(onClick, event.currentTarget.checked ? 0 : undefined)} />{t('After:')}</label>
        <input class="seconds" type="number" min="0" step="0.01" aria-label={t('Advance after (seconds)')} disabled={!timingEditable || afterMs === undefined} value={((afterMs ?? 0) / MS_PER_SECOND).toFixed(2)} onchange={(event) => { const value = event.currentTarget.valueAsNumber; if (Number.isFinite(value) && value >= 0) changeTiming(onClick, Math.round(value * MS_PER_SECOND)); }} />
      </div>
    </div>
    <input bind:this={soundFile} class="file" type="file" accept=".wav,audio/wav" tabindex="-1" aria-label={t('Other Sound…')} onchange={(event) => { void chooseSoundFile(event.currentTarget.files?.[0]); event.currentTarget.value = ''; }} />
    <button class="big" style:--w="38px" disabled={!timingEditable} onclick={() => changeTiming(onClick, afterMs, getSlides(doc.pres))}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="7" y="3" width="14" height="11" rx="1" /><rect x="3" y="8" width="14" height="11" rx="1" /><path d="m14 17 3 2-3 2" /></svg>
      <span>{t('Apply To All')}</span>
    </button>
  </section>
</div>

<style>
  /* Measured from Mac PowerPoint 16 (see POWERPOINT_PARITY.md "Native
     geometry audit"): a 72 pt row, groups split by a rule with 10 pt on each
     side, large buttons as wide as PowerPoint's own, 26 pt control rows on a
     32 pt pitch starting 4 pt down. */
  .transitions { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
  .cluster { display: flex; flex: none; align-items: stretch; padding: 0 10px; border-right: 1px solid var(--ok-border); }
  .cluster:first-child { padding-left: 4px; }
  .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled, label:has(input:disabled) { opacity: 0.4; cursor: default; }
  button[aria-expanded='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: var(--w); padding: 4px 1px; font-size: 11px; line-height: 1.15; text-align: center; }
  .big > span:last-child { max-width: var(--w); margin: 0 -2px; }
  /* Japanese labels wrap per character, so they get at least six characters
     a line and a smaller size that fits three lines (PowerPoint widens them). */
  .big > span:last-child:lang(ja) { max-width: max(calc(var(--w) - 4px), 6em); font-size: 10px; line-height: 1.1; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 11px; }
  .big svg { width: 32px; height: 32px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.1; }
  .anchor { position: relative; display: flex; }
  .timing { display: grid; grid-template-columns: auto auto; grid-auto-rows: 26px; gap: 6px 15px; align-self: flex-start; margin: 4px 15px 0 0; font-size: 12px; }
  .field { justify-self: end; }
  label, .check-row { display: flex; align-items: center; gap: 6px; white-space: nowrap; }
  input[type='checkbox'] { margin: 0; accent-color: var(--ok-accent); }
  .field svg { width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.1; }
  select { box-sizing: border-box; width: 102px; height: 26px; font: inherit; padding: 1px 4px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-panel); color: var(--ok-text); }
  .seconds { box-sizing: border-box; width: 78px; height: 24px; font: inherit; padding: 1px 4px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-panel); color: var(--ok-text); }
  .file { position: absolute; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
  .art { width: 48px; height: 32px; fill: none; stroke: currentColor; stroke-width: 1; }
  .slide { fill: var(--ok-panel); stroke: currentColor; }
  .next { fill: color-mix(in srgb, var(--ok-accent) 35%, transparent); stroke: none; }
  .next.half { fill: color-mix(in srgb, var(--ok-accent) 18%, transparent); }
  /* PowerPoint's Effect Options menu: 58 pt rows with a thumbnail. */
  .menu { position: fixed; z-index: 300; display: flex; flex-direction: column; width: 210px; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; align-items: center; gap: 8px; height: 58px; padding: 0 8px 0 4px; border: none; border-radius: 0; font-size: 13px; text-align: left; }
  .menu .check { width: 12px; visibility: hidden; }
  .menu button[aria-checked='true'] .check { visibility: visible; }
  .menu .thumb { width: 74px; height: 50px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.2; }
</style>
