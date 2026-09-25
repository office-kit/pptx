<script lang="ts">
  // Plays the slide in hand the way the presentation surface would, over the
  // editor rather than in it.
  //
  // It draws its own copy of the slide from the document as it stands, so an
  // effect changed a second ago plays as it is now — nothing here waits for a
  // save to land or reads the preview's saved state. Playing changes nothing:
  // no transaction, no history entry, no dirty flag. The copy, the player's
  // timers and the keyboard go back to the editor when it closes.
  import { onMount } from 'svelte';
  import type { SlideAnimationStep } from '@office-kit/pptx';
  import {
    createAnimationPlayer,
    type AnimationPlayer,
    type UnsupportedReason,
  } from '../core/animation-player.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  interface Props {
    readonly svg: string;
    readonly steps: readonly SlideAnimationStep[];
    readonly onclose: () => void;
  }
  const { svg, steps, onclose }: Props = $props();

  const REASONS: Record<UnsupportedReason, string> = {
    notInMainSeq: 'Runs on its own trigger, not the slide’s clicks',
    notModelled: 'Uses an effect, target or start this library only reads',
    unknownTiming: 'Starts when an effect of unstated length ends',
    ambiguousTarget: 'Names an id more than one object on the slide carries',
  };

  let stage = $state<HTMLDivElement>();
  let dialog = $state<HTMLDialogElement>();
  let player = $state<AnimationPlayer>();
  let progress = $state({ cursor: 0, stops: 0 });
  let skipped = $state<{ reason: UnsupportedReason; count: number }[]>([]);

  // A modal dialog so the editor behind it takes neither clicks nor keys while
  // the slide is playing, and the focus goes back where it was on close.
  onMount(() => {
    const opener = document.activeElement;
    dialog?.showModal();
    // Opening it focuses the first control, which would then eat the arrow keys
    // the show is driven with. The dialog itself takes the focus instead.
    dialog?.focus();
    return () => {
      if (opener instanceof HTMLElement) opener.focus();
    };
  });

  // Keyed on what it is playing: a slide edited, undone or switched while the
  // show is open would leave a copy of a deck that no longer exists on screen.
  $effect(() => {
    const root = stage;
    if (!root) return;
    root.innerHTML = svg;
    // The stage is the slide's own shape, not a guess: a 4:3 or portrait deck
    // is shown at its own ratio rather than letterboxed into a widescreen box.
    const box = root.querySelector('svg')?.viewBox.baseVal;
    if (box !== undefined && box.width > 0 && box.height > 0)
      root.style.setProperty('--ok-slide-ratio', String(box.width / box.height));
    const created = createAnimationPlayer({
      root,
      steps,
      reducedMotion: () => matchMedia('(prefers-reduced-motion: reduce)').matches,
      onChange: (current) => {
        progress = { cursor: current.cursor, stops: current.stopCount };
      },
    });
    const counts = new Map<UnsupportedReason, number>();
    for (const entry of created.unsupported)
      counts.set(entry.reason, (counts.get(entry.reason) ?? 0) + 1);
    skipped = [...counts].map(([reason, count]) => ({ reason, count }));
    player = created;
    created.reset();
    return () => {
      created.dispose();
      player = undefined;
      root.innerHTML = '';
      root.style.removeProperty('--ok-slide-ratio');
    };
  });

  // The show owns the keyboard while it is open. A modal dialog keeps the focus
  // but not the events: a key pressed in it still bubbles to the editor's own
  // window handler, where Ctrl+K, Ctrl+S or an undo would act on a document the
  // viewer cannot see. So every key stops here — and the ones the browser
  // handles itself, on a focused control, keep their default behaviour.
  function onkeydown(event: KeyboardEvent): void {
    event.stopPropagation();
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    const target = event.target instanceof Element ? event.target : null;
    // A control uses the keys it owns — Space and Enter press a button, a field
    // takes everything — and the show keeps the rest.
    if (target?.closest('select, input, textarea') != null) return;
    const onButton = target?.closest('button') != null;
    if (onButton && (event.key === ' ' || event.key === 'Enter')) return;
    const forward = ['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(event.key);
    const back = ['ArrowLeft', 'ArrowUp', 'PageUp'].includes(event.key);
    if (!forward && !back) return;
    event.preventDefault();
    event.stopPropagation();
    if (forward) player?.advance();
    else player?.back();
  }
</script>

<dialog
  bind:this={dialog}
  tabindex="-1"
  aria-label={t('Play animations')}
  {onkeydown}
  onclose={onclose}
>
  <div class="slide" bind:this={stage} aria-label={t('Animation preview')}></div>
  <div class="controls">
    <span aria-live="polite">{t('Click')} {progress.cursor} / {progress.stops}</span>
    <button class="ok-btn" disabled={progress.cursor <= 0} onclick={() => player?.back()}
      >{t('Previous')}</button
    >
    <button
      class="ok-btn"
      disabled={progress.cursor >= progress.stops}
      onclick={() => player?.advance()}>{t('Next')}</button
    >
    <button class="ok-btn" onclick={() => player?.reset()}>{t('Restart')}</button>
    <button class="ok-btn" onclick={() => dialog?.close()}>{t('Stop')}</button>
  </div>
  {#if skipped.length > 0}
    <ul class="skipped">
      {#each skipped as entry (entry.reason)}
        <li>{entry.count} × {t(REASONS[entry.reason])}</li>
      {/each}
    </ul>
  {/if}
</dialog>

<style>
  dialog {
    width: 100vw;
    max-width: 100vw;
    height: 100vh;
    max-height: 100vh;
    margin: 0;
    border: 0;
    padding: 20px;
    display: grid;
    align-content: center;
    justify-items: center;
    gap: 12px;
    background: #0d0f18;
    color: white;
  }
  dialog::backdrop {
    background: #0d0f18;
  }
  .slide {
    width: min(90vw, calc((100vh - 160px) * var(--ok-slide-ratio, 1.7778)));
    aspect-ratio: var(--ok-slide-ratio, 1.7778);
    background: white;
    box-shadow: 0 20px 60px #0008;
  }
  .slide :global(svg) {
    display: block;
    width: 100%;
    height: 100%;
  }
  .controls {
    display: flex;
    gap: 8px;
    align-items: center;
    font-size: 12px;
  }
  .skipped {
    margin: 0;
    padding: 0;
    list-style: none;
    color: #ffd9a8;
    font-size: 11px;
    text-align: center;
  }
</style>
