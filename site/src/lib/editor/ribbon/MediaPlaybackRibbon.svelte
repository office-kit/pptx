<script lang="ts">
  import { tick } from 'svelte';
  import MediaTrimDialog from '../ui/MediaTrimDialog.svelte';
  import {
    getShapeMedia,
    getShapeMediaPlayback,
    setShapeMediaPlayback,
    getShapeId,
    type MediaPlayback,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { getMediaPreview } from '../core/media-preview.svelte.ts';

  // Mac PowerPoint writes numSld=999 when Play Across Slides is enabled.
  const acrossSlidesCount = 999;
  const volumeOptions = [
    { label: 'Low', value: 0.2 },
    { label: 'Medium', value: 0.5 },
    { label: 'High', value: 0.8 },
  ] as const;
  const editor = getEditor();
  const doc = editor.doc;
  const preview = getMediaPreview(editor);
  let error = $state('');
  let trimOpen = $state(false);
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

  function apply(label: string, options: Partial<MediaPlayback>): boolean {
    const target = selected;
    if (target === null) return false;
    try {
      doc.transact(t(label), () => setShapeMediaPlayback(target.shape, options));
      error = '';
      return true;
    } catch (cause) {
      error = cause instanceof Error ? cause.message : String(cause);
      return false;
    }
  }

  $effect(() => {
    if (trimOpen && selected) preview.command('pause', getShapeId(selected.shape));
  });

  const currentBookmark = $derived(selected ? preview.bookmarkAtCurrent(selected.playback) : null);
  const selectedBookmark = $derived(selected && preview.state.shapeId === getShapeId(selected.shape) ? preview.state.bookmarkIndex : null);
  const sortedBookmarks = $derived(
    selected
      ? [...(selected.playback.bookmarks ?? [])]
          .map((bookmark, originalIndex) => ({ bookmark, originalIndex }))
          .sort((a, b) => a.bookmark.timeMs - b.bookmark.timeMs)
      : [],
  );
  const playLabel = $derived(preview.state.playing ? 'Pause' : 'Play');
  function addBookmark(): void {
    if (!selected || selected.media.kind === 'online' || currentBookmark !== null) return;
    const bookmarks = [...(selected.playback.bookmarks ?? [])];
    const timeMs = preview.state.currentTime * 1000;
    const used = new Set(bookmarks.map(bookmark => bookmark.name));
    let name = `Bookmark ${bookmarks.length + 1}`;
    let suffix = bookmarks.length + 1;
    while (used.has(name)) name = `Bookmark ${suffix += 1}`;
    bookmarks.push({ name, timeMs });
    bookmarks.sort((a, b) => a.timeMs - b.timeMs);
    if (apply('Add Bookmark', { bookmarks })) {
      const added = bookmarks.find(bookmark => bookmark.name === name);
      preview.selectBookmark(bookmarks.findIndex(bookmark => bookmark.name === name), added?.timeMs);
    }
  }
  function removeBookmark(): void {
    if (!selected || selectedBookmark === null) return;
    const bookmarks = [...(selected.playback.bookmarks ?? [])];
    const target = sortedBookmarks[selectedBookmark];
    if (!target) return;
    bookmarks.splice(target.originalIndex, 1);
    if (!apply('Remove Bookmark', { bookmarks })) return;
    preview.selectBookmark(null);
  }

  function mediaFileName(media: Extract<NonNullable<ReturnType<typeof getShapeMedia>>, { kind: 'audio' | 'video' }>): string {
    const name = media.partName.split('/').pop() || `${media.kind}1`;
    if (/\.[a-z0-9]+$/i.test(name)) return name;
    const extensionByType: Record<string, string> = {
      'audio/mpeg': 'mp3',
      'audio/mp4': 'm4a',
      'audio/wav': 'wav',
      'audio/x-wav': 'wav',
      'video/quicktime': 'mov',
      'video/webm': 'webm',
      'video/mp4': 'mp4',
    };
    const extension = extensionByType[media.contentType.split(';', 1)[0]!.toLowerCase()];
    return extension ? `${name}.${extension}` : name;
  }

  function saveMediaAs(): void {
    const target = selected;
    if (!target || target.media.kind === 'online') return;
    const url = URL.createObjectURL(new Blob([new Uint8Array(target.media.bytes)], { type: target.media.contentType }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = mediaFileName(target.media);
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
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

  function changeFade(input: HTMLInputElement, edge: 'inMs' | 'outMs'): void {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) return;
    const fade = selected?.playback.fade ?? { inMs: 0, outMs: 0 };
    apply(edge === 'inMs' ? 'Fade In' : 'Fade Out', {
      fade: { ...fade, [edge]: input.valueAsNumber * 1000 },
    });
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
  {#if trimOpen}<MediaTrimDialog shape={selected.shape} onclose={() => trimOpen = false} />{/if}
  <div class="group">
    <div class="items">
      <button class="action" type="button" aria-label={t(playLabel)} disabled={selected.media.kind === 'online'} onclick={() => preview.command(preview.state.playing ? 'pause' : 'play', getShapeId(selected.shape))}>{t(playLabel)}</button>
      <button class="action" type="button" aria-label={t('Save Media As')} disabled={selected.media.kind === 'online'} onclick={saveMediaAs}>{t('Save Media As')}</button>
      <button class="action" type="button" aria-label={t('Add Bookmark')} disabled={selected.media.kind === 'online' || currentBookmark !== null} onclick={addBookmark}>{t('Add Bookmark')}</button>
      <button class="action" type="button" aria-label={t('Remove Bookmark')} disabled={selectedBookmark === null} onclick={removeBookmark}>{t('Remove Bookmark')}</button>
    </div>
    <span class="title">{t('Bookmarks')}</span>
  </div>
  <div class="group"><div class="items"><button class="action" disabled={selected.media.kind === 'online'} onclick={() => trimOpen = true}>{t(selected.media.kind === 'video' ? 'Trim Video' : 'Trim Audio')}</button></div><span class="title">{t('Editing')}</span></div>
  <div class="group">
    <div class="items">
      {#each [{ key: 'inMs', label: 'Fade In' }, { key: 'outMs', label: 'Fade Out' }] as edge}
        <label class="field number-field">
          <span>{t(edge.label)}</span>
          <span class="number"><input class="ok-input" type="number" min="0" step="0.05" required aria-label={t(edge.label)} value={(selected.playback.fade?.[edge.key as 'inMs' | 'outMs'] ?? 0) / 1000} onchange={event => changeFade(event.currentTarget, edge.key as 'inMs' | 'outMs')} /><span>s</span></span>
        </label>
      {/each}
    </div>
    <span class="title">{t('Fade Duration')}</span>
  </div>
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
    {#if error || preview.state.error}<span class="error" role="alert">{error || preview.state.error}</span>{/if}
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
