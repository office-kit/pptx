<script lang="ts">
  import type { Snippet } from 'svelte';
  import { t } from '../i18n/i18n.svelte.ts';

  type Props = {
    duration: number;
    position: number;
    start?: number;
    end?: number;
    fadeIn?: number;
    fadeOut?: number;
    onseek?: (position: number) => void;
    children?: Snippet;
  };

  const STEP_MS = 50;
  let {
    duration,
    position,
    start = $bindable(0),
    end = $bindable(0),
    fadeIn = $bindable(0),
    fadeOut = $bindable(0),
    onseek,
    children,
  }: Props = $props();

  const clamp = (value: number, minimum: number, maximum: number) => Math.max(minimum, Math.min(maximum, value));
  const trackPercent = (value: number) => duration > 0 ? `${clamp(value, 0, duration) / duration * 100}%` : '0%';
  const trackValue = (event: Event) => (event.currentTarget as HTMLInputElement).valueAsNumber;
  // Short clips still need a valid range even though normal keyboard moves are 50 ms.
  const minimumRange = $derived(Math.min(STEP_MS, Math.max(1, duration)));
  const length = $derived(Math.max(minimumRange, end - start));
  const fadeInPosition = $derived(start + clamp(fadeIn, 0, length));
  const fadeOutPosition = $derived(end - clamp(fadeOut, 0, length));

  function seek(value: number) {
    const next = clamp(value, start, end);
    onseek?.(next);
  }

  function changeStart(value: number) {
    const next = clamp(value, 0, Math.max(0, end - minimumRange));
    start = next;
    fadeIn = clamp(fadeIn, 0, end - next);
    fadeOut = clamp(fadeOut, 0, end - next);
    seek(position);
  }

  function changeEnd(value: number) {
    const next = clamp(value, Math.min(duration, start + minimumRange), duration);
    end = next;
    fadeIn = clamp(fadeIn, 0, next - start);
    fadeOut = clamp(fadeOut, 0, next - start);
    seek(position);
  }

  function changeFadeIn(value: number) {
    fadeIn = clamp(value - start, 0, end - start);
  }

  function changeFadeOut(value: number) {
    fadeOut = clamp(end - value, 0, end - start);
  }
</script>

<div class="timeline" style={`--start:${trackPercent(start)};--end:${trackPercent(end)};--fade-in:${trackPercent(fadeInPosition)};--fade-out:${trackPercent(fadeOutPosition)};--position:${trackPercent(position)}`}>
  <div class="track">
    {@render children?.()}
    <span aria-hidden="true" class="trimmed before"></span>
    <span aria-hidden="true" class="trimmed after"></span>
    <span aria-hidden="true" class="selected"></span>
    <span aria-hidden="true" class="fade-zone fade-in"></span>
    <span aria-hidden="true" class="fade-zone fade-out"></span>
    <span aria-hidden="true" class="trim-handle start"></span>
    <span aria-hidden="true" class="trim-handle end"></span>
    <span aria-hidden="true" class="fade-handle in"></span>
    <span aria-hidden="true" class="fade-handle out"></span>
    <span aria-hidden="true" class="playhead"></span>
  </div>
  <input class="axis current" aria-label={t('Current Position')} type="range" min="0" max={duration} step={STEP_MS} value={position} disabled={duration <= 0 || end <= start} oninput={event => seek(trackValue(event))} />
  <input class="axis trim start" aria-label={t('Start Trim')} type="range" min="0" max={duration} step={STEP_MS} value={start} disabled={duration <= 0} oninput={event => changeStart(trackValue(event))} />
  <input class="axis trim end" aria-label={t('End Trim')} type="range" min="0" max={duration} step={STEP_MS} value={end} disabled={duration <= 0} oninput={event => changeEnd(trackValue(event))} />
  <input class="axis fade in" aria-label={t('Fade In')} type="range" min="0" max={length} step={STEP_MS} value={fadeIn} disabled={duration <= 0} oninput={event => changeFadeIn(trackValue(event) + start)} />
  <input class="axis fade out" aria-label={t('Fade Out')} type="range" min="0" max={length} step={STEP_MS} value={fadeOut} disabled={duration <= 0} oninput={event => changeFadeOut(end - trackValue(event))} />
</div>

<style>
  .timeline { position: relative; flex: 1; height: 72px; min-width: 180px; margin: 5px 0; }
  .track { position: absolute; inset: 25px 0 25px; overflow: visible; border-radius: 3px; background: #444; }
  .track > :global(*) { position: absolute; inset: 0; width: 100%; height: 100%; }
  .trimmed { background: rgb(0 0 0 / 45%); }
  .trimmed.before { right: auto; width: var(--start); }
  .trimmed.after { left: var(--end); width: calc(100% - var(--end)); }
  .selected { left: var(--start); width: calc(var(--end) - var(--start)); background: #b48a1f88; }
  .fade-zone { top: 0; height: 100%; background: #fff4; }
  .fade-zone.fade-in { left: var(--start); width: calc(var(--fade-in) - var(--start)); }
  .fade-zone.fade-out { left: var(--fade-out); width: calc(var(--end) - var(--fade-out)); }
  .trim-handle, .fade-handle, .playhead { top: -9px; bottom: -9px; width: 2px; height: auto; transform: translateX(-1px); }
  .trim-handle { background: #ffd21f; }
  .trim-handle.start { left: var(--start); }
  .trim-handle.end { left: var(--end); }
  .fade-handle { background: white; box-shadow: 0 0 0 1px #222; }
  .fade-handle.in { left: var(--fade-in); }
  .fade-handle.out { left: var(--fade-out); }
  .playhead { left: var(--position); width: 2px; transform: translateX(-1px); background: #238bff; z-index: 2; }
  .axis { position: absolute; left: -7px; width: calc(100% + 14px); height: 20px; margin: 0; appearance: none; background: transparent; pointer-events: none; }
  .axis::-webkit-slider-runnable-track { height: 20px; background: transparent; }
  .axis::-moz-range-track { height: 20px; background: transparent; }
  .axis::-webkit-slider-thumb { width: 14px; height: 24px; margin-top: -2px; border: 0; border-radius: 3px; background: transparent; appearance: none; pointer-events: auto; cursor: ew-resize; }
  .axis::-moz-range-thumb { width: 14px; height: 24px; border: 0; border-radius: 3px; background: transparent; pointer-events: auto; cursor: ew-resize; }
  .axis.current { top: 26px; z-index: 5; }
  .axis.current::-webkit-slider-thumb { width: 14px; background: #238bff; }
  .axis.current::-moz-range-thumb { width: 14px; background: #238bff; }
  .axis.trim { top: 4px; z-index: 6; }
  .axis.trim::-webkit-slider-thumb { background: #ffd21f; }
  .axis.trim::-moz-range-thumb { background: #ffd21f; }
  .axis.fade { top: 49px; left: calc(var(--start) - 7px); width: calc(var(--end) - var(--start) + 14px); z-index: 7; }
  .axis.fade.out { direction: rtl; }
  .axis.fade::-webkit-slider-thumb { width: 14px; background: white; box-shadow: 0 0 0 1px #222; }
  .axis.fade::-moz-range-thumb { width: 14px; background: white; box-shadow: 0 0 0 1px #222; }
  .axis:focus-visible::-webkit-slider-thumb { outline: 2px solid #72b6ff; outline-offset: 2px; }
  .axis:focus-visible::-moz-range-thumb { outline: 2px solid #72b6ff; outline-offset: 2px; }
  .axis:focus-visible { outline: none; }
</style>
