<script lang="ts">
  import { onMount } from 'svelte';
  import { t } from '../i18n/i18n.svelte.ts';

  const { bytes }: { bytes: Uint8Array } = $props();
  const columns = 512;
  let path = $state('');
  let unavailable = $state(false);

  onMount(() => {
    let disposed = false;
    // Offline decoding does not start an audio output device or request playback.
    const context = new OfflineAudioContext(1, 1, 8000);
    void context.decodeAudioData(new Uint8Array(bytes).buffer).then(buffer => {
      if (disposed) return;
      const peaks = new Float32Array(columns);
      for (let channel = 0; channel < buffer.numberOfChannels; channel++) {
        const samples = buffer.getChannelData(channel);
        for (let i = 0; i < samples.length; i++) {
          const column = Math.min(columns - 1, Math.floor(i * columns / samples.length));
          peaks[column] = Math.max(peaks[column]!, Math.abs(samples[i]!));
        }
      }
      path = Array.from(peaks, (peak, x) => {
        const height = Math.min(1, peak) * 22;
        return `M${x} ${24 - height}V${24 + height}`;
      }).join('');
    }, () => {
      // Some playable video containers cannot be decoded by Web Audio.
      if (!disposed) unavailable = true;
    });
    return () => { disposed = true; };
  });
</script>

<div class="waveform" title={unavailable ? t('Waveform unavailable') : undefined}>
  <svg viewBox="0 0 512 48" preserveAspectRatio="none" aria-label={t('Audio waveform')} role="img">
    <path d={path || 'M0 24H512'} fill="none" stroke="currentColor" stroke-width="1" />
  </svg>
  {#if unavailable}<span>{t('Waveform unavailable')}</span>{/if}
</div>

<style>
  .waveform { position: absolute; inset: 0; pointer-events: none; color: #b7c1cc; }
  svg { display: block; width: 100%; height: 100%; }
  span { position: absolute; inset: 0; display: grid; place-items: center; font-size: 10px; background: #202020b0; color: #ccc; }
</style>
