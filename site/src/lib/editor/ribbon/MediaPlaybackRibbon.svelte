<script lang="ts">
  import { tick } from 'svelte';
  import {
    getShapeMedia,
    getShapeMediaPlayback,
    setShapeMediaPlayback,
    type MediaPlayback,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  // Mac PowerPoint writes numSld=999 when Play Across Slides is enabled.
  const acrossSlidesCount = 999;
  const volumeOptions = [
    { label: 'Low', value: 0.2 },
    { label: 'Medium', value: 0.5 },
    { label: 'High', value: 0.8 },
  ] as const;
  const editor = getEditor();
  const doc = editor.doc;
  let error = $state('');
  let volumeOpen = $state(false);
  let volumeTrigger = $state<HTMLButtonElement>();
  let volumeMenu = $state<HTMLDivElement>();

  const selected = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    if (shapes.length !== 1) return null;
    const shape = shapes[0]!;
    const media = getShapeMedia(shape);
    const playback = getShapeMediaPlayback(shape);
    return media !== null && playback !== null ? { shape, media, playback } : null;
  });

  const delaySeconds = $derived(selected?.playback.delayMs === undefined ? 0 : selected.playback.delayMs / 1000);

  function apply(label: string, options: Partial<MediaPlayback>): void {
    const target = selected;
    if (target === null) return;
    try {
      doc.transact(t(label), () => setShapeMediaPlayback(target.shape, options));
      error = '';
    } catch (cause) {
      error = cause instanceof Error ? cause.message : String(cause);
    }
  }

  function changeDelay(input: HTMLInputElement): void {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) return;
    const milliseconds = Math.round(input.valueAsNumber * 1000);
    if (!Number.isSafeInteger(milliseconds)) {
      error = t('Delay must be a whole number of milliseconds');
      return;
    }
    apply('Start delay', { delayMs: milliseconds });
  }

  function closeVolume(restore = true): void {
    volumeOpen = false;
    if (restore) volumeTrigger?.focus();
  }

  async function showVolume(): Promise<void> {
    volumeOpen = !volumeOpen;
    if (volumeOpen) {
      await tick();
      (volumeMenu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? volumeMenu?.querySelector<HTMLButtonElement>('button'))?.focus();
    }
  }

  function placeVolumeMenu(node: HTMLElement): void {
    const bounds = volumeTrigger!.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8))}px`;
  }

  function volumeKeys(event: KeyboardEvent): void {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); closeVolume(); return; }
    if (event.key === 'Tab') { closeVolume(false); return; }
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const items = [...volumeMenu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length]?.focus();
  }
</script>

<svelte:window onpointerdown={event => { if (volumeOpen && !volumeMenu?.contains(event.target as Node) && !volumeTrigger?.contains(event.target as Node)) closeVolume(false); }} onblur={() => { if (volumeOpen) closeVolume(false); }} onresize={() => { if (volumeOpen) closeVolume(false); }} />

{#if selected}
  <div class="group">
    <div class="items">
      <label class="field">
        <span>{t('Start')}</span>
        <select aria-label={t('Start')} value={selected.playback.autoplay ? 'automatic' : 'click'} onchange={(event) => apply('Start', { autoplay: event.currentTarget.value === 'automatic' })}>
          <option value="automatic">{t('Automatically')}</option>
          <option value="click">{t('When Clicked On')}</option>
        </select>
      </label>
      <label class="check">
        <input type="checkbox" checked={selected.playback.loop} onchange={(event) => apply('Loop until stopped', { loop: event.currentTarget.checked })} />
        <span>{t('Loop until stopped')}</span>
      </label>
      <div class="volume-menu">
        <button class="volume-trigger" bind:this={volumeTrigger} type="button" aria-label={t('Volume')} title={t('Volume')} aria-haspopup="menu" aria-expanded={volumeOpen} onclick={showVolume}>{t('Volume')} ▾</button>
        {#if volumeOpen}
          <div class="menu" role="menu" aria-label={t('Volume')} tabindex="-1" bind:this={volumeMenu} use:placeVolumeMenu onkeydown={volumeKeys}>
            {#each volumeOptions as option}
              <button role="menuitemradio" aria-checked={!selected.playback.muted && selected.playback.volume === option.value} onclick={() => { apply(option.label, { volume: option.value, muted: false }); closeVolume(); }}><span class="check" aria-hidden="true">{!selected.playback.muted && selected.playback.volume === option.value ? '✓' : ''}</span>{t(option.label)}</button>
            {/each}
            <button role="menuitemradio" aria-checked={selected.playback.muted} onclick={() => { apply('Mute', { muted: true }); closeVolume(); }}><span class="check" aria-hidden="true">{selected.playback.muted ? '✓' : ''}</span>{t('Mute')}</button>
          </div>
        {/if}
      </div>
      {#if selected.media.kind === 'audio'}
        <button class="action" type="button" aria-label={t('Play in Background')} onclick={() => apply('Play in Background', { autoplay: true, slideCount: acrossSlidesCount, loop: true, hideWhenStopped: true })}>{t('Play in Background')}</button>
        <label class="check">
          <input type="checkbox" checked={(selected.playback.slideCount ?? 1) > 1} onchange={(event) => apply('Play Across Slides', { slideCount: event.currentTarget.checked ? acrossSlidesCount : 1 })} />
          <span>{t('Play Across Slides')}</span>
        </label>
      {/if}
      {#if selected.media.kind === 'video'}
        <label class="check">
          <input type="checkbox" checked={selected.playback.fullScreen} onchange={(event) => apply('Play Full Screen', { fullScreen: event.currentTarget.checked })} />
          <span>{t('Play Full Screen')}</span>
        </label>
      {/if}
      <label class="check">
        <input type="checkbox" checked={selected.playback.rewindAfterPlaying ?? false} onchange={(event) => apply('Rewind After Playing', { rewindAfterPlaying: event.currentTarget.checked })} />
        <span>{t('Rewind After Playing')}</span>
      </label>
      <label class="check">
        <input type="checkbox" checked={selected.playback.hideWhenStopped} onchange={(event) => apply('Hide when not playing', { hideWhenStopped: event.currentTarget.checked })} />
        <span>{t('Hide when not playing')}</span>
      </label>
    </div>
    {#if error}<span class="error" role="alert">{error}</span>{/if}
    <span class="title">{t('Playback')}</span>
  </div>
  {#if selected.playback.autoplay}
    <div class="group timing">
      <div class="items">
        <label class="field number-field">
          <span>{t('Start delay (seconds)')}</span>
          <span class="number"><input class="ok-input" type="number" min="0" step="0.001" required aria-label={t('Start delay (seconds)')} value={delaySeconds} onchange={(event) => changeDelay(event.currentTarget)} /><span>s</span></span>
        </label>
      </div>
      <span class="title">{t('Timing')}</span>
    </div>
  {/if}
{/if}

<style>
  .group { display: flex; flex-direction: column; justify-content: space-between; flex-shrink: 0; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  .timing { min-width: 145px; }
  .items { display: flex; align-items: center; flex: 1; gap: 8px; }
  .field, .check { display: flex; align-items: center; gap: 4px; color: var(--ok-text); font: inherit; font-size: 11px; white-space: nowrap; }
  .field { flex-direction: column; align-items: stretch; gap: 2px; }
  select, input { font: inherit; color: inherit; }
  select { min-width: 112px; padding: 3px 4px; }
  .number { display: flex; align-items: center; gap: 3px; }
  .number input { width: 52px; min-width: 0; padding: 2px 4px; }
  .volume-menu { position: relative; display: flex; align-items: center; }
  .volume-trigger { padding: 3px 6px; border: 1px solid var(--ok-border); border-radius: 3px; background: var(--ok-surface, transparent); color: var(--ok-text); font: inherit; font-size: 11px; white-space: nowrap; }
  .menu { position: fixed; z-index: 400; min-width: 130px; padding: 4px; background: var(--ok-panel); color: var(--ok-text); border: 1px solid var(--ok-border); border-radius: 5px; box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; gap: 6px; width: 100%; border: 0; padding: 5px 8px; background: transparent; color: inherit; text-align: left; font: inherit; font-size: 12px; }
  .menu button:hover, .menu button:focus-visible { background: var(--ok-accent); color: white; outline: none; }
  .menu .check { width: 14px; }
  .check { max-width: 125px; }
  .action { padding: 3px 6px; border: 1px solid var(--ok-border); border-radius: 3px; background: var(--ok-surface, transparent); color: var(--ok-text); font: inherit; font-size: 11px; white-space: nowrap; }
  .check input { margin: 0; }
  .check span { white-space: normal; }
  .error { color: var(--ok-danger, #b42318); font-size: 11px; padding-top: 3px; }
  .title { text-align: center; font-size: 10px; color: var(--ok-text-2); padding: 4px 0 1px; }
</style>
