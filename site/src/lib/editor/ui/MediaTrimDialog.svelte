<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getShapeMedia, getShapeMediaPlayback, setShapeMediaPlayback, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import MediaTrimTimeline from './MediaTrimTimeline.svelte';
  import MediaWaveform from './MediaWaveform.svelte';

  const { shape, onclose }: { shape: SlideShapeData; onclose: () => void } = $props();
  const editor = getEditor();
  const media = untrack(() => getShapeMedia(shape));
  const playback = untrack(() => getShapeMediaPlayback(shape));
  let dialog: HTMLDialogElement;
  let player = $state<HTMLMediaElement>()!;
  let url = $state('');
  let duration = $state(0);
  let start = $state(0);
  let end = $state(0);
  let fadeIn = $state(0);
  let fadeOut = $state(0);
  let position = $state(0);
  let paused = $state(true);
  let volume = $state(0.5);
  let error = $state('');
  const progressIntervalMs = 25;
  let progressTimer: number | undefined;
  const length = $derived(end - start);
  const valid = $derived(duration > 0 && start >= 0 && end <= duration && length > 0);
  const time = (ms: number) => `${Math.floor(ms / 60000)}:${(Math.max(0, ms) / 1000 % 60).toFixed(3).padStart(6, '0')}`;

  onMount(() => {
    dialog.showModal();
    if (!media || media.kind === 'online') { error = t('The media could not be loaded'); return; }
    url = URL.createObjectURL(new Blob([new Uint8Array(media.bytes)], { type: media.contentType }));
    const handlePlay = () => {
      paused = false;
      scheduleProgress();
    };
    const handlePause = () => {
      paused = true;
      cancelProgress();
      progress();
    };
    player.addEventListener('play', handlePlay);
    player.addEventListener('pause', handlePause);
    return () => {
      cancelProgress();
      player.removeEventListener('play', handlePlay);
      player.removeEventListener('pause', handlePause);
      player.pause();
      URL.revokeObjectURL(url);
    };
  });
  function loaded() {
    if (!Number.isFinite(player.duration)) { error = t('The media could not be loaded'); return; }
    error = '';
    duration = player.duration * 1000;
    end = duration - (playback?.trim?.endMs ?? 0);
    start = playback?.trim?.startMs ?? 0;
    fadeIn = playback?.fade?.inMs ?? 0;
    fadeOut = playback?.fade?.outMs ?? 0;
    seek(start);
  }
  function seek(value: number) {
    if (!duration) return;
    position = Math.max(start, Math.min(end, value));
    player.currentTime = position / 1000;
    updateVolume();
  }
  function updateVolume() {
    const incoming = fadeIn > 0 ? (position - start) / fadeIn : 1;
    const outgoing = fadeOut > 0 ? (end - position) / fadeOut : 1;
    player.volume = volume * Math.max(0, Math.min(1, incoming, outgoing));
  }
  function cancelProgress() {
    if (progressTimer === undefined) return;
    window.clearTimeout(progressTimer);
    progressTimer = undefined;
  }
  function scheduleProgress() {
    cancelProgress();
    const tick = () => {
      progressTimer = undefined;
      progress();
      if (!player.paused && position < end) progressTimer = window.setTimeout(tick, progressIntervalMs);
    };
    progressTimer = window.setTimeout(tick, 0);
  }
  function progress() {
    const current = player.currentTime * 1000;
    position = Math.max(start, Math.min(end, current));
    if (current >= end) {
      player.pause();
      if (Math.abs(current - end) > 0.001) player.currentTime = end / 1000;
      position = end;
      cancelProgress();
    }
    updateVolume();
  }
  async function play() {
    if (!paused) { player.pause(); return; }
    if (position >= end || position < start) seek(start);
    try { await player.play(); error = ''; }
    catch { error = t('Unable to play media · Retry'); }
  }
  function apply() {
    if (!valid) return;
    try {
      editor.doc.transact(t('Trim'), () => setShapeMediaPlayback(shape, {
        trim: { startMs: start, endMs: duration - end },
        fade: { inMs: fadeIn, outMs: fadeOut },
      }));
      onclose();
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
</script>

<dialog bind:this={dialog} oncancel={onclose} aria-label={t(media?.kind === 'video' ? 'Trim Video' : 'Trim Audio')}>
  <header><strong>{t(media?.kind === 'video' ? 'Trim Video' : 'Trim Audio')}</strong><span>{t('Duration')}: {time(Math.max(0, length))}</span></header>
  <div class="preview">
    {#if media?.kind === 'video'}
      <!-- svelte-ignore a11y_media_has_caption -->
      <video bind:this={player} src={url} onloadedmetadata={loaded} ontimeupdate={progress} onplay={() => paused = false} onpause={() => paused = true} onerror={() => error = t('The media could not be loaded')}></video>
    {:else}
      <svg class="speaker" viewBox="0 0 100 100" aria-hidden="true"><path d="M12 38h20L52 20v60L32 62H12Z" fill="currentColor"/><path d="M64 33q20 17 0 34M76 21q33 29 0 58" fill="none" stroke="currentColor" stroke-width="7" stroke-linecap="round"/></svg>
      <audio bind:this={player} src={url} onloadedmetadata={loaded} ontimeupdate={progress} onplay={() => paused = false} onpause={() => paused = true} onerror={() => error = t('The media could not be loaded')}></audio>
    {/if}
  </div>
  <div class="scrubber">
    <span>{time(position)}</span>
    <MediaTrimTimeline bind:start bind:end bind:fadeIn bind:fadeOut {duration} {position} onseek={seek}>
      {#if media && media.kind !== 'online'}<MediaWaveform bytes={media.bytes} />{/if}
    </MediaTrimTimeline>
    <span>−{time(Math.max(0, end - position))}</span>
  </div>
  <div class="transport"><button aria-label={t('Nudge Backward')} disabled={!valid} onclick={() => seek(position - 50)}>◀|</button><button aria-label={t(paused ? 'Play' : 'Pause')} disabled={!valid} onclick={play}>{paused ? '▶' : 'Ⅱ'}</button><button aria-label={t('Nudge Forward')} disabled={!valid} onclick={() => seek(position + 50)}>|</button></div>
  {#if error}<p role="alert">{error}</p>{/if}
  <footer><input aria-label={t('Volume')} type="range" min="0" max="1" step="0.01" bind:value={volume} oninput={updateVolume} /><div><button onclick={onclose}>{t('Cancel')}</button><button class="primary" disabled={!valid} onclick={apply}>{t('Trim')}</button></div></footer>
</dialog>

<style>
  dialog { width: min(800px, 90vw); border: 1px solid #444; border-radius: 16px; padding: 20px; background: #202020; color: #eee; }
  dialog::backdrop { background: #0008; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 16px; }
  header { padding-bottom: 16px; border-bottom: 1px solid #444; }
  .preview { margin: 18px 0; height: min(42vh, 360px); background: black; display: grid; place-items: center; }
  video { width: 100%; height: 100%; object-fit: contain; }
  .speaker { width: 100px; height: 100px; color: #ddd; }
  .scrubber { display: flex; align-items: center; gap: 14px; font-variant-numeric: tabular-nums; }
  .transport { display: flex; justify-content: center; gap: 12px; margin: 12px; }
  button { padding: 6px 16px; background: #353535; color: white; border: 0; border-radius: 6px; font: inherit; }
  button:disabled { opacity: .4; }
  .primary { background: #bd4e08; }
  footer { border-top: 1px solid #444; padding-top: 16px; }
  footer div { display: flex; gap: 12px; }
  footer > input { width: 150px; }
  p { color: #ffb8b8; }
</style>
