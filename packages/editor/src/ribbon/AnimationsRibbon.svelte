<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // The reference desktop app's (Mac, 16) Animations tab: Preview | Entrance Effects gallery |
  // Emphasis Effects gallery | Exit Effects | Path Animation | Effect Options,
  // Animation Pane, Trigger, Animation Painter | Start and Duration. In a
  // 1512 pt window both galleries show five 64 pt tiles; at 1200 pt the
  // Emphasis gallery collapses into a button and Entrance shows six.
  import { placeBelowTrigger } from './place-menu.ts';
  import {
    getSlideAnimations,
    removeSlideAnimation,
    setShapeAnimation,
    updateSlideAnimation,
    type AnimationEffect,
    type AnimationPatch,
    type AnimationStartCondition,
  } from '@office-kit/pptx';
  import { tick, untrack } from 'svelte';
  import AnimationPlayback from '../ui/AnimationPlayback.svelte';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import RibbonGallery from './RibbonGallery.svelte';
  import EffectGalleryPopover from './EffectGalleryPopover.svelte';
  import AnimationColorPalette from './AnimationColorPalette.svelte';
  import {
    ALL_TILES,
    EMPHASIS_TILES,
    ENTRANCE_TILES,
    EXIT_TILES,
    SEQUENCE,
    effectOptionSections,
    optionChecked,
    optionLabel,
    stepOptions,
    takesColor,
    takesSequence,
    type EffectTile,
  } from './animation-gallery.ts';

  const STARTS: ReadonlyArray<readonly [AnimationStartCondition, string]> = [
    ['click', 'On Click'],
    ['withPrevious', 'With Previous'],
    ['afterPrevious', 'After Previous'],
  ];
  const MS_PER_SECOND = 1000;
  const TILE_WIDTH = 64;
  // The command row width (CSS px) the reference desktop app's expanded layout needs; below
  // it the Emphasis gallery becomes a button and Entrance gains a tile.
  const EXPANDED_FROM = 1380;

  const editor = getEditor();
  const doc = editor.doc;
  const shapeId = $derived(selectedShapeId(doc.selection));
  // The first effect on the selected shape is the one the galleries show and
  // replace, like the reference desktop app's single-effect galleries.
  const step = $derived.by(() => {
    doc.version;
    const slide = doc.currentSlide;
    if (!slide || shapeId === null) return null;
    return getSlideAnimations(slide).find((item) => item.id !== null && item.targetShapeIds.includes(shapeId)) ?? null;
  });
  const optionSections = $derived(effectOptionSections(step?.effect ?? null));
  const checkedTile = $derived(ALL_TILES.find((item) => item.effect === step?.effect) ?? null);
  // Keys are unique within a gallery, not across them: Blinds is both an
  // entrance and an exit.
  const checkedIn = (kind: EffectTile['kind']) => (checkedTile?.kind === kind ? checkedTile.key : null);
  let width = $state(typeof window === 'undefined' ? Number.POSITIVE_INFINITY : window.innerWidth);
  // The threshold fits the reference desktop app's English labels; longer ones (Japanese)
  // collapse the Emphasis gallery whenever the expanded row does not fit.
  let squeezed = $state(false);
  let row = $state<HTMLDivElement>();
  const expanded = $derived(width >= EXPANDED_FROM && !squeezed);
  $effect.pre(() => {
    width;
    getLocale();
    squeezed = false;
  });
  $effect(() => {
    const last = row?.lastElementChild;
    if (!expanded || !row || !last) return;
    if (last.getBoundingClientRect().right > row.getBoundingClientRect().right) squeezed = true;
  });
  const label = (item: { en: string; ja: string }) => (getLocale() === 'ja' ? item.ja : item.en);

  // A new preset replaces the effect and takes the preset's own duration, as
  // the reference desktop app's galleries do; the library resets it when the preset changes.
  function choose(effect: AnimationEffect) {
    const slide = doc.currentSlide;
    const shape = editor.selectedShapes()[0];
    if (!slide || !shape) return;
    doc.transact(t('Animation'), () => {
      if (step?.id != null) updateSlideAnimation(slide, step.id, { effect });
      else setShapeAnimation(shape, { effect });
    });
  }
  function patch(value: AnimationPatch) {
    const slide = doc.currentSlide;
    if (!slide || step?.id == null) return;
    const id = step.id;
    try {
      doc.transact(t('Animation'), () => updateSlideAnimation(slide, id, value));
    } catch (error) {
      editor.toast('error', error instanceof Error ? error.message : String(error));
    }
  }

  type Menu = 'emphasis' | 'exit' | 'options';
  let menu = $state<Menu | null>(null);
  const toggle = (name: Menu) => (menu = menu === name ? null : name);
  function chooseFromPopover(item: EffectTile) {
    menu = null;
    choose(item.effect);
  }

  // Preview plays the current slide's effects over the editor.
  const steps = $derived.by(() => { doc.version; return doc.currentSlide ? [...getSlideAnimations(doc.currentSlide)] : []; });
  let previewing = $state(false);

  // Animation Painter: the next object selected takes the selected one's effect.
  let painter = $state<{ from: number; options: Parameters<typeof setShapeAnimation>[1] } | null>(null);
  function startPainter() {
    if (!step?.effect || shapeId === null) return;
    painter = {
      from: shapeId,
      options: {
        effect: step.effect,
        ...stepOptions(step),
        ...(step.build !== 'custom' ? { build: step.build } : {}),
        ...(step.durationMs != null ? { durationMs: step.durationMs } : {}),
        ...(step.delayMs != null ? { delayMs: step.delayMs } : {}),
        ...(step.start !== 'unknown' ? { start: step.start } : {}),
      },
    };
  }
  $effect(() => {
    const id = shapeId;
    untrack(() => {
      if (!painter || id === null || id === painter.from) return;
      const options = painter.options;
      painter = null;
      const slide = doc.currentSlide;
      const target = editor.selectedShapes()[0];
      if (!slide || !target) return;
      doc.transact(t('Animation Painter'), () => {
        for (const item of getSlideAnimations(slide)) if (item.id !== null && item.targetShapeIds.includes(id)) removeSlideAnimation(slide, item.id);
        setShapeAnimation(target, options);
      });
    });
  });

  async function showPane() {
    // It sits with the shape's Size & Properties sections in the Format pane.
    editor.showShapeFormat('size');
    await tick();
    const pane = editor.shell?.querySelector<HTMLElement>('section[data-animation-pane]');
    pane?.scrollIntoView({ block: 'nearest' });
    pane?.focus();
  }
</script>

<svelte:window onkeydown={(event) => { if (event.key === 'Escape') { painter = null; menu = null; } }} onpointerdown={(event) => { if (menu && !(eventTarget(event) as Element).closest?.('.anchor')) menu = null; }} />

{#snippet star(kind: EffectTile['kind'])}
  <svg class="star {kind}" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.4 6.7 19.4l1.2-6L3.4 9.3l6-.7z" /></svg>
{/snippet}
{#snippet effectPopover(name: Menu, items: readonly EffectTile[])}
  <EffectGalleryPopover label={t(name === 'exit' ? 'Exit Effects' : 'Emphasis Effects')} tiles={items} checked={checkedIn(name === 'exit' ? 'exit' : 'emphasis')} choose={chooseFromPopover}>
    {#snippet tile(item)}{@render star(item.kind)}{/snippet}
  </EffectGalleryPopover>
{/snippet}

<div class="animations" lang={getLocale()} bind:this={row} bind:clientWidth={width}>
  <section class="cluster" role="group" aria-label={t('Preview')}>
    <button class="big" style:--w="50px" disabled={steps.length === 0} onclick={() => (previewing = true)}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.4 6.7 19.4l1.2-6L3.4 9.3l6-.7z" /><path d="m15 15 6 3.5-6 3.5z" fill="currentColor" /></svg>
      <span>{t('Preview')}</span>
    </button>
  </section>
  {#if previewing && doc.currentSlide}<AnimationPlayback svg={doc.currentSvg} {steps} onclose={() => (previewing = false)} />{/if}
  <section class="cluster" role="group" aria-label={t('Entrance Effects')}>
    <RibbonGallery label={t('Entrance Effects')} items={ENTRANCE_TILES} checked={checkedIn('entrance')} visible={expanded ? 5 : 6} tileWidth={TILE_WIDTH} disabled={shapeId === null} name={label} choose={(item) => choose(item.effect)}>
      {#snippet tile(item)}{@render star(item.kind)}{/snippet}
    </RibbonGallery>
  </section>
  <section class="cluster" role="group" aria-label={t('Emphasis Effects')}>
    {#if expanded}
      <RibbonGallery label={t('Emphasis Effects')} items={EMPHASIS_TILES} checked={checkedIn('emphasis')} visible={5} tileWidth={TILE_WIDTH} disabled={shapeId === null} name={label} choose={(item) => choose(item.effect)}>
        {#snippet tile(item)}{@render star(item.kind)}{/snippet}
      </RibbonGallery>
    {:else}
      <div class="anchor">
        <button class="big" style:--w="55px" aria-haspopup="dialog" aria-expanded={menu === 'emphasis'} disabled={shapeId === null} onclick={() => toggle('emphasis')}>
          <span class="icon-row">{@render star('emphasis')}<span class="arrow" aria-hidden="true">⌄</span></span>
          <span>{t('Emphasis Effects')}</span>
        </button>
        {#if menu === 'emphasis'}{@render effectPopover('emphasis', EMPHASIS_TILES)}{/if}
      </div>
    {/if}
  </section>
  <section class="cluster" role="group" aria-label={t('Exit Effects')}>
    <div class="anchor">
      <button class="big" style:--w="50px" aria-haspopup="dialog" aria-expanded={menu === 'exit'} disabled={shapeId === null} onclick={() => toggle('exit')}>
        <span class="icon-row">{@render star('exit')}<span class="arrow" aria-hidden="true">⌄</span></span>
        <span>{t('Exit Effects')}</span>
      </button>
      {#if menu === 'exit'}{@render effectPopover('exit', EXIT_TILES)}{/if}
    </div>
  </section>
  <section class="cluster" role="group" aria-label={t('Path Animation')}>
    <button class="big" style:--w="57px" aria-haspopup="menu" disabled title={t('Motion paths are not supported by the library yet.')}>
      <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h8M4 12h12M4 18h6" /><path d="M14 18c4 0 6-3 6-6" stroke-dasharray="2 2" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
      <span>{t('Path Animation')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Advanced Animation')}>
    <div class="anchor">
      <button class="big" style:--w="50px" aria-haspopup="menu" aria-expanded={menu === 'options'} disabled={step?.id == null || (optionSections.length === 0 && !takesColor(step.effect) && !takesSequence(step.effect))} onclick={() => toggle('options')}>
        <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 3l2 4.4 4.7.5-3.5 3.2.9 4.6L10 13.4l-4.1 2.3.9-4.6-3.5-3.2 4.7-.5z" /><circle cx="17" cy="17" r="3" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
        <span>{t('Effect Options')}</span>
      </button>
      {#if menu === 'options' && step}
        <!-- The reference desktop app's Effect Options menu: the effect's own sections
             (Direction, Shapes, Spokes, Amount), the colour palette of a
             colour effect, then Sequence. -->
        <div class="menu options" role="menu" use:placeBelowTrigger aria-label={t('Effect Options')}>
          {#each optionSections as group, at (`${at}:${group.heading}`)}
            {#if at > 0}<hr />{/if}
            <div class="heading" role="presentation">{t(group.heading)}</div>
            {#each group.items as item (item.en)}
              {@const exit = step.presetClass === 'exit'}
              <button role="menuitemradio" aria-checked={optionChecked(step, item)} onclick={() => { menu = null; patch(item.patch); }}>
                <span class="check" aria-hidden="true">✓</span>
                {#if item.arrow !== undefined}<svg class="dir" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 12h16m-5-5 5 5-5 5" transform="rotate({-(item.arrow + (exit ? 180 : 0))} 12 12)" /></svg>{/if}
                {label(optionLabel(item, exit))}
              </button>
            {/each}
          {/each}
          {#if takesColor(step.effect)}
            {#if optionSections.length > 0}<hr />{/if}
            <AnimationColorPalette value={step.color} choose={(color) => { menu = null; patch({ color }); }} />
          {/if}
          {#if takesSequence(step.effect)}
            {#if optionSections.length > 0 || takesColor(step.effect)}<hr />{/if}
            <div class="heading" role="presentation">{t('Sequence')}</div>
            {#each SEQUENCE as item (item.build)}
              <button role="menuitemradio" aria-checked={step.build === item.build} onclick={() => { menu = null; patch({ build: item.build }); }}><span class="check" aria-hidden="true">✓</span>{t(item.en)}</button>
            {/each}
          {/if}
        </div>
      {/if}
    </div>
    <button class="big" style:--w="57px" onclick={showPane}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.4 6.7 19.4l1.2-6L3.4 9.3l6-.7z" /><path d="M16 15h6M16 18h6M16 21h6" /></svg>
      <span>{t('Animation Pane')}</span>
    </button>
    <button class="big" style:--w="50px" aria-haspopup="menu" disabled title={t('Triggers are not supported by the library yet.')}>
      <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M13 2 5 13h6l-2 9 9-12h-6z" fill="#d9a520" stroke="#a87c10" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
      <span>{t('Trigger')}</span>
    </button>
    <button class="big" style:--w="57px" aria-pressed={painter !== null} disabled={!step?.effect} onclick={() => (painter ? (painter = null) : startPainter())}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 3l2 4.4 4.7.5-3.5 3.2.9 4.6L10 13.4l-4.1 2.3.9-4.6-3.5-3.2 4.7-.5z" /><path d="m15 21 6-6-2-2-6 6z" /></svg>
      <span>{t('Animation Painter')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Timing')}>
    <div class="timing">
      <label>
        <svg viewBox="0 0 16 16" aria-hidden="true"><path d="m5 3 8 5-8 5z" /></svg>
        <span>{t('Start:')}</span>
        <select aria-label={t('Start')} disabled={step?.id == null} value={step?.start ?? ''} onchange={(event) => patch({ start: event.currentTarget.value as AnimationStartCondition })}>
          {#if step?.id == null}<option value=""></option>{/if}
          {#each STARTS as [value, name] (value)}<option {value}>{t(name)}</option>{/each}
        </select>
      </label>
      <label>
        <svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="6.5" /><path d="M8 4v4l3 2" /></svg>
        <span>{t('Duration:')}</span>
        <input type="number" min="0.01" step="0.25" aria-label={t('Duration')} disabled={step?.id == null} value={step?.durationMs != null ? (step.durationMs / MS_PER_SECOND).toFixed(2) : ''} onchange={(event) => { const value = event.currentTarget.valueAsNumber; if (Number.isFinite(value) && value > 0) patch({ durationMs: Math.round(value * MS_PER_SECOND) }); }} />
      </label>
    </div>
  </section>
</div>

<style>
  /* Geometry measured from the reference desktop app (Mac, 16) (NATIVE_PARITY.md, "Native
     geometry audit"): a 72 pt row, a rule with 10 pt each side between groups,
     the reference desktop app's own large-button widths and 26 pt rows on a 32 pt pitch. */
  .animations { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
  .cluster { display: flex; flex: none; align-items: stretch; padding: 0 10px; border-right: 1px solid var(--ok-border); }
  .cluster:first-child { padding-left: 4px; }
  .cluster:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled, label:has(:disabled) { opacity: 0.4; cursor: default; }
  button[aria-pressed='true'], button[aria-expanded='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: var(--w); padding: 4px 1px; font-size: 11px; line-height: 1.15; text-align: center; }
  .big > span:last-child { max-width: var(--w); margin: 0 -2px; }
  /* Japanese labels wrap per character, so they get at least six characters
     a line and a smaller size that fits three lines (the reference desktop app widens them). */
  .big > span:last-child:lang(ja) { max-width: max(calc(var(--w) - 4px), 6em); font-size: 10px; line-height: 1.1; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 11px; }
  svg { width: 32px; height: 32px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.1; }
  .star.entrance path { fill: #3f9b4a; stroke: #2d7a37; }
  .star.emphasis path { fill: #d9a520; stroke: #a87c10; }
  .star.exit path { fill: #c8423b; stroke: #9c2d27; }
  .anchor { position: relative; display: flex; }
  .menu { position: fixed; z-index: 300; display: flex; flex-direction: column; min-width: 200px; max-height: 70vh; overflow-y: auto; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; flex: none; align-items: center; gap: 8px; height: 32px; padding: 0 12px 0 6px; border: none; border-radius: 0; font-size: 13px; text-align: left; }
  .menu svg { width: 20px; height: 20px; }
  .menu .dir { stroke: var(--ok-accent); }
  .menu .heading { display: flex; align-items: center; height: 23px; padding: 0 12px; font-size: 12px; color: var(--ok-text-2); }
  .menu hr { width: 100%; margin: 6px 0; border: none; border-top: 1px solid var(--ok-border); }
  .menu .check { width: 12px; visibility: hidden; }
  .menu button[aria-checked='true'] .check { visibility: visible; }
  .options { width: 214px; }
  .timing { display: flex; flex-direction: column; align-items: flex-end; gap: 6px; margin-top: 4px; font-size: 12px; }
  .timing label { display: flex; align-items: center; gap: 6px; height: 26px; white-space: nowrap; }
  .timing svg { width: 16px; height: 16px; }
  select, input { box-sizing: border-box; font: inherit; padding: 1px 4px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-panel); color: var(--ok-text); }
  select { width: 102px; height: 26px; }
  input { width: 78px; height: 24px; }
</style>
