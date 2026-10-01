<script lang="ts">
  import { untrack } from 'svelte';
  import { getShapeId, setShapeMediaPlayback, type MediaPlayback, type ShapeMedia, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { getMediaPreview } from '../core/media-preview.svelte.ts';

  let { shape, media, playback }: { shape: SlideShapeData; media: ShapeMedia; playback: MediaPlayback } = $props();
  const editor = getEditor();
  const preview = getMediaPreview(editor);
  let player = $state<HTMLMediaElement>();
  const shapeId = $derived(getShapeId(shape));
  const bookmarks = $derived([...(playback.bookmarks ?? [])].sort((a, b) => a.timeMs - b.timeMs));
  const selectedBookmark = $derived(preview.state.shapeId === shapeId ? preview.state.bookmarkIndex : null);
  const time = (seconds: number) => {
    const total = Math.max(0, Math.floor(seconds * 100) / 100);
    const minutes = Math.floor(total / 60);
    return `${minutes}:${(total - minutes * 60).toFixed(2).padStart(5, '0')}`;
  };
  const previewTickMs = 25;
  const seekStepSeconds = 0.25;
  const mediaBytes = $derived(media.kind === 'online' ? null : media.bytes);
  const contentType = $derived(media.kind === 'online' ? '' : media.contentType);
  const mediaUrl = $derived(media.kind === 'online' ? media.url : '');
  let src = $state('');
  $effect(() => {
    const next = mediaBytes
      ? URL.createObjectURL(new Blob([new Uint8Array(mediaBytes)], { type: contentType }))
      : mediaUrl;
    src = next;
    return () => { if (mediaBytes) URL.revokeObjectURL(next); };
  });
  $effect(() => {
    const target = player;
    const id = shapeId;
    src;
    if (target) return untrack(() => preview.attach(id, target));
  });
  $effect(() => {
    if (!preview.state.playing || preview.state.shapeId !== shapeId) return;
    const timer = window.setInterval(timeUpdate, previewTickMs);
    return () => window.clearInterval(timer);
  });
  $effect(() => {
    playback.volume;
    playback.muted;
    playback.fade;
    if (player) untrack(applyVolume);
  });
  function positionControls(node: HTMLDivElement) {
    const area = node.closest<HTMLElement>('.canvas-area');
    const host = node.parentElement;
    if (!area || !host) return;
    const minimumToolbarWidth = 380;
    const toolbarGap = 4;
    let frame = 0;
    // Mac PowerPoint keeps playback controls horizontal below the rotated media bounds.
    // Render in canvas coordinates so group scaling and rotation affect only the media.
    area.append(node);
    const position = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => {
        const bounds = area.getBoundingClientRect();
        const anchor = host.getBoundingClientRect();
        const width = Math.min(area.clientWidth, Math.max(minimumToolbarWidth, anchor.width));
        const centered = (anchor.left + anchor.right - width) / 2 - bounds.left;
        const left = Math.max(0, Math.min(area.clientWidth - width, centered));
        node.style.width = `${width}px`;
        node.style.left = `${left + area.scrollLeft}px`;
        node.style.top = `${anchor.bottom - bounds.top + area.scrollTop + toolbarGap}px`;
      });
    };
    const resize = new ResizeObserver(position);
    const mutation = new MutationObserver(position);
    // The shape moves via inline styles; canvas resize and scrolling also move its toolbar.
    for (let ancestor: HTMLElement | null = host; ancestor; ancestor = ancestor.parentElement) {
      resize.observe(ancestor);
      mutation.observe(ancestor, { attributes: true, attributeFilter: ['style'] });
      if (ancestor === area) break;
    }
    area.addEventListener('scroll', position);
    position();
    return { destroy() {
      cancelAnimationFrame(frame);
      resize.disconnect();
      mutation.disconnect();
      area.removeEventListener('scroll', position);
      node.remove();
    } };
  }
  function bookmarkSeek(index: number) {
    const bookmark = bookmarks[index];
    if (!bookmark) return;
    seek(bookmark.timeMs / 1000);
    preview.selectBookmark(index, bookmark.timeMs);
  }
  const trimStart = $derived((playback.trim?.startMs ?? 0) / 1000);
  const trimEnd = $derived(Math.max(trimStart, preview.state.duration - (playback.trim?.endMs ?? 0) / 1000));
  function loaded() {
    if (!player) return;
    player.currentTime = trimStart;
    player.volume = playback.volume;
    player.muted = playback.muted;
  }
  function seek(time: number) { preview.command('seek', shapeId, Math.min(trimEnd, Math.max(trimStart, time))); }
  function step(seconds: number) {
    const current = preview.state.shapeId === shapeId ? preview.state.currentTime : trimStart;
    seek(current + seconds);
  }
  function applyVolume() {
    if (!player) return;
    const elapsed = Math.max(0, player.currentTime - trimStart) * 1000;
    const remaining = Math.max(0, trimEnd - player.currentTime) * 1000;
    const fade = playback.fade;
    const fadeIn = fade?.inMs ? Math.min(1, elapsed / fade.inMs) : 1;
    const fadeOut = fade?.outMs ? Math.min(1, remaining / fade.outMs) : 1;
    player.volume = playback.volume * Math.min(fadeIn, fadeOut);
    player.muted = playback.muted;
  }
  function beginPlay() {
    if (!player) return;
    if (player.currentTime >= trimEnd || player.currentTime < trimStart) player.currentTime = trimStart;
    applyVolume();
  }
  function finishPlay() {
    if (!player || trimEnd <= trimStart) return;
    if (playback.loop) {
      preview.command('seek', shapeId, trimStart);
      preview.command('play', shapeId);
    } else {
      preview.command('pause', shapeId);
      preview.command('seek', shapeId, playback.rewindAfterPlaying ? trimStart : trimEnd);
    }
  }
  function timeUpdate() {
    if (!player) return;
    applyVolume();
    if (!player.paused && trimEnd > trimStart && player.currentTime >= trimEnd) finishPlay();
  }
  function toggleMute() {
    const muted = !playback.muted;
    if (player) player.muted = muted;
    editor.doc.transact(t(muted ? 'Mute' : 'Unmute'), () => setShapeMediaPlayback(shape, { muted }));
  }
  function togglePlay() {
    if (preview.state.playing) {
      preview.command('pause', shapeId);
      return;
    }
    if (preview.state.shapeId === shapeId && preview.state.currentTime >= trimEnd) {
      preview.command('seek', shapeId, trimStart);
    }
    preview.command('play', shapeId);
  }
</script>

<!-- svelte-ignore a11y_click_events_have_key_events -->
<!-- svelte-ignore a11y_no_static_element_interactions -->
<div class="media-preview" role="presentation" onclick={event => event.stopPropagation()} onpointerdown={event => event.stopPropagation()} onkeydown={event => event.stopPropagation()}>
  {#if media.kind === 'video'}
    <!-- svelte-ignore a11y_media_has_caption -->
    <video style:visibility={preview.state.showVideoFrame ? undefined : 'hidden'} bind:this={player} src={src} preload="metadata" aria-label={t('Media preview')} onloadedmetadata={loaded} ontimeupdate={timeUpdate} onplay={beginPlay} onended={finishPlay}></video>
  {:else}
    <audio bind:this={player} src={src} preload="metadata" onloadedmetadata={loaded} ontimeupdate={timeUpdate} onplay={beginPlay} onended={finishPlay}></audio>
  {/if}
  <div class="controls" use:positionControls role="presentation" aria-label={t('Media controls')} onclick={event => event.stopPropagation()} onpointerdown={event => event.stopPropagation()} onkeydown={event => event.stopPropagation()}>
    <button type="button" aria-label={preview.state.playing ? t('Pause') : t('Play')} onclick={togglePlay}><svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true">{#if preview.state.playing}<path fill="currentColor" d="M5 3h5v18H5zm9 0h5v18h-5z" />{:else}<path fill="currentColor" d="M5 2v20l16-10z" />{/if}</svg></button>
    <div class="seekbar" style={`--played:${trimEnd > trimStart ? Math.min(100, Math.max(0, (preview.state.currentTime - trimStart) / (trimEnd - trimStart) * 100)) : 0}%`}>
      <input aria-label={t('Media position')} type="range" min={trimStart} max={trimEnd || 0} step="any" value={preview.state.shapeId === shapeId ? preview.state.currentTime : trimStart} oninput={event => seek(event.currentTarget.valueAsNumber)} />
      {#each bookmarks as bookmark, index}
        {@const markerTime = Math.min(trimEnd, Math.max(trimStart, bookmark.timeMs / 1000))}
        <button type="button" class:selected={selectedBookmark === index} class="bookmark" aria-label={bookmark.name} title={bookmark.name} style={`left:${trimEnd > trimStart ? ((markerTime - trimStart) / (trimEnd - trimStart)) * 100 : 0}%`} onclick={() => bookmarkSeek(index)}>●</button>
      {/each}
    </div>
    <button type="button" aria-label={t('Back')} onclick={() => step(-seekStepSeconds)}><svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><path fill="currentColor" d="M17 5h2v14h-2zM15 5v14L4 12z" /></svg></button>
    <button type="button" aria-label={t('Forward')} onclick={() => step(seekStepSeconds)}><svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><path fill="currentColor" d="M4 5h2v14H4zm3 0v14l11-7z" /></svg></button>
    <span class="time">{time(preview.state.shapeId === shapeId ? preview.state.currentTime : 0)}</span>
    <button type="button" class:muted={playback.muted} aria-label={t(playback.muted ? 'Unmute' : 'Mute')} onclick={toggleMute}><svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="currentColor" d="M3 9h4l5-4v14l-5-4H3z" /><path d="M15 8a6 6 0 0 1 0 8m3-11a10 10 0 0 1 0 14" fill="none" stroke="currentColor" stroke-width="1.6" /></svg></button>
  </div>
  {#if preview.state.error}<span class="error" role="alert">{preview.state.error}</span>{/if}
</div>

<style>
  .media-preview { position:absolute; inset:0; pointer-events:none; }
  video { width:100%; height:100%; object-fit:contain; pointer-events:none; }
  audio { display:none; }
  .controls { position:absolute; z-index:1; height:38px; box-sizing:border-box; display:flex; align-items:center; gap:8px; padding:4px 8px; background:linear-gradient(#fff, #eff0f2); color:#303030; border:1px solid #c8c9cc; border-radius:4px; box-shadow:0 1px 2px #0001; pointer-events:auto; font-size:12px; }
  .controls button { display:flex; align-items:center; justify-content:center; flex-shrink:0; color:inherit; background:transparent; border:0; padding:1px 3px; cursor:pointer; }
  .controls button:disabled { opacity:.4; cursor:default; }
  .controls button.muted { color:#efc451; }
  .controls button:focus-visible { outline:2px solid #376cb5; outline-offset:1px; }
  .seekbar { position:relative; flex:1; min-width:40px; display:flex; align-items:center; }
  .seekbar input[type=range] { appearance:none; width:100%; height:28px; margin:0; border:1px solid #bfc0c3; border-radius:0; background:linear-gradient(to right, #c2c0c0 0 var(--played), #f8f9fa var(--played) 100%); cursor:pointer; }
  .seekbar input[type=range]::-webkit-slider-thumb { appearance:none; width:1px; height:26px; background:#999; }
  .seekbar input[type=range]::-moz-range-thumb { width:1px; height:26px; border:0; border-radius:0; background:#999; }
  .time { min-width:48px; font-variant-numeric:tabular-nums; }
  .bookmark { position:absolute; top:50%; transform:translate(-50%, -50%); color:white !important; text-shadow:0 0 2px #333; font-size:10px; }
  .bookmark.selected { color:#f3c546 !important; }
  .error { position:absolute; right:4px; bottom:-60px; color:#b42318; background:white; padding:2px 4px; }
</style>
