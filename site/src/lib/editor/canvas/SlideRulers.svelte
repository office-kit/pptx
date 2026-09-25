<script lang="ts">
  import { t } from '../i18n/i18n.svelte.ts';
  let { area, stage, zoom }: { area: HTMLElement; stage: HTMLElement; zoom: number } = $props();
  let root: HTMLDivElement;
  let bounds = $state({ x: 0, y: 0, width: 0, height: 0, slideWidth: 0, slideHeight: 0 });
  const thickness = 22;
  const pixelsPerCentimeter = $derived(96 * zoom / 2.54);
  // Keep subdivisions legible when zoomed out without allocating off-screen ticks.
  const interval = $derived(pixelsPerCentimeter >= 24 ? 0.25 : pixelsPerCentimeter >= 12 ? 0.5 : pixelsPerCentimeter >= 6 ? 1 : 2);
  const labelInterval = $derived(pixelsPerCentimeter >= 24 ? 1 : pixelsPerCentimeter >= 12 ? 2 : pixelsPerCentimeter >= 6 ? 5 : 10);
  function ticks(offset: number, length: number, viewport: number) {
    const center = offset + length / 2;
    const first = Math.ceil((Math.max(thickness, offset) - center) / pixelsPerCentimeter / interval);
    const last = Math.floor((Math.min(viewport, offset + length) - center) / pixelsPerCentimeter / interval);
    return Array.from({ length: Math.max(0, last - first + 1) }, (_, i) => {
      const value = (first + i) * interval;
      return { value, position: center + value * pixelsPerCentimeter, major: Number.isInteger(value), labeled: value % labelInterval === 0 };
    });
  }
  const horizontal = $derived(ticks(bounds.x, bounds.slideWidth, bounds.width));
  const vertical = $derived(ticks(bounds.y, bounds.slideHeight, bounds.height));
  $effect(() => {
    function measure() {
      const origin = root.getBoundingClientRect();
      const slide = stage.getBoundingClientRect();
      bounds = { x: slide.left - origin.left, y: slide.top - origin.top, width: area.clientWidth + thickness, height: area.clientHeight + thickness, slideWidth: slide.width, slideHeight: slide.height };
    }
    const observer = new ResizeObserver(measure);
    observer.observe(area);
    observer.observe(stage);
    area.addEventListener('scroll', measure, { passive: true });
    measure();
    return () => { observer.disconnect(); area.removeEventListener('scroll', measure); };
  });
</script>

<div class="rulers" bind:this={root}>
  <svg class="horizontal" role="img" aria-label={t('Horizontal ruler (centimeters)')} width={bounds.width} height={thickness}>
    <rect x={bounds.x} y="0" width={bounds.slideWidth} height={thickness} />
    {#each horizontal as mark}
      <line data-value={mark.value} x1={mark.position} x2={mark.position} y1={mark.major ? 16 : 18} y2={thickness} />
      {#if mark.labeled}<text x={mark.position} y="12" text-anchor="middle">{Math.abs(mark.value)}</text>{/if}
    {/each}
  </svg>
  <svg class="vertical" role="img" aria-label={t('Vertical ruler (centimeters)')} width={thickness} height={bounds.height}>
    <rect x="0" y={bounds.y} width={thickness} height={bounds.slideHeight} />
    {#each vertical as mark}
      <line data-value={mark.value} y1={mark.position} y2={mark.position} x1={mark.major ? 16 : 18} x2={thickness} />
      {#if mark.labeled}<text transform={`translate(9 ${mark.position}) rotate(-90)`} text-anchor="middle">{Math.abs(mark.value)}</text>{/if}
    {/each}
  </svg>
  <div class="corner"></div>
</div>

<style>
  .rulers { position: absolute; inset: 0; overflow: hidden; pointer-events: none; z-index: 2; }
  svg { position: absolute; left: 0; top: 0; background: var(--ok-canvas-bg); overflow: hidden; }
  rect, .corner { fill: var(--ok-panel); background: var(--ok-panel); }
  line { stroke: var(--ok-text-2); stroke-width: 1; }
  text { fill: var(--ok-text-2); font: 10px sans-serif; }
  .corner { position: absolute; left: 0; top: 0; width: 22px; height: 22px; }
</style>
