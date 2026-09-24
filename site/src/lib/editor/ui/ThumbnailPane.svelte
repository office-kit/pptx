<script lang="ts">
  import { onMount } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import SlideNavigator from './SlideNavigator.svelte';
  const editor = getEditor();
  const MIN_WIDTH = 90;
  const MAX_WIDTH = 400;
  let pane: HTMLDivElement;
  let width = $state(200);
  let drag: { id: number; x: number; width: number } | null = null;
  function setWidth(value: number) {
    editor.thumbnailWidth = Math.max(MIN_WIDTH, Math.min(MAX_WIDTH, pane.parentElement!.clientWidth / 3, value));
  }
  function start(event: PointerEvent) {
    if (event.button !== 0) return;
    event.preventDefault();
    const handle = event.currentTarget as HTMLElement;
    handle.focus();
    drag = { id: event.pointerId, x: event.clientX, width: pane.clientWidth };
    handle.setPointerCapture(event.pointerId);
  }
  function move(event: PointerEvent) {
    if (drag?.id === event.pointerId) setWidth(drag.width + event.clientX - drag.x);
  }
  function keys(event: KeyboardEvent) {
    if (event.key === 'Escape' && drag) { event.preventDefault(); event.stopPropagation(); editor.thumbnailWidth = drag.width; drag = null; return; }
    if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    setWidth(event.key === 'Home' ? MIN_WIDTH : event.key === 'End' ? MAX_WIDTH : pane.clientWidth + (event.key === 'ArrowRight' ? 10 : -10));
  }
  onMount(() => {
    const observer = new ResizeObserver(() => width = pane.clientWidth);
    observer.observe(pane);
    return () => observer.disconnect();
  });
</script>

<div class="thumbnail-pane" bind:this={pane}>
  <SlideNavigator />
  <!-- A focusable separator implements the ARIA window-splitter pattern. -->
  <!-- svelte-ignore a11y_no_noninteractive_tabindex, a11y_no_noninteractive_element_interactions -->
  <div class="resize" role="separator" tabindex="0" aria-label={t('Thumbnail pane width')} aria-orientation="vertical" aria-valuemin={MIN_WIDTH} aria-valuemax={MAX_WIDTH} aria-valuenow={width} onpointerdown={start} onpointermove={move} onpointerup={() => drag = null} onlostpointercapture={() => drag = null} onpointercancel={() => { if (drag) editor.thumbnailWidth = drag.width; drag = null; }} onkeydown={keys}></div>
</div>

<style>
  .thumbnail-pane { position: relative; display: grid; min-height: 0; min-width: 0; }
  .resize { position: absolute; top: 0; right: -3px; bottom: 0; width: 6px; cursor: ew-resize; touch-action: none; z-index: 2; }
  .resize:focus-visible { outline: 2px solid var(--ok-accent); }
</style>
