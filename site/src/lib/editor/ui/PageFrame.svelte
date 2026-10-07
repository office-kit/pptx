<script lang="ts">
  // The portrait page of Notes Page view and the handout and notes masters:
  // 7.5 × 10 in, fitted to the window with 22 pt around it (Mac PowerPoint
  // opens these views at 96% in a 1512 × 900 pt window), or at the zoom chosen.
  // Contents are placed in page fractions; `--page-scale` is screen px per point.
  import type { Snippet } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { PAGE_HEIGHT_PT, PAGE_WIDTH_PT } from '../core/master-geometry.ts';

  let { label, children }: { label: string; children: Snippet } = $props();
  const editor = getEditor();
  const FIT_MARGIN = 22;
  let areaWidth = $state(0);
  let areaHeight = $state(0);
  $effect(() => {
    const fit = Math.min((areaWidth - 2 * FIT_MARGIN) / PAGE_WIDTH_PT, (areaHeight - 2 * FIT_MARGIN) / PAGE_HEIGHT_PT);
    editor.pageFitZoom = fit > 0 ? fit : 1;
    if (editor.pageAutoFitZoom) editor.pageZoom = editor.pageFitZoom;
  });
</script>

<div class="page-area ok-scroll" role="region" aria-label={label} bind:clientWidth={areaWidth} bind:clientHeight={areaHeight}>
  <div class="page" style:width="{PAGE_WIDTH_PT * editor.pageZoom}px" style:height="{PAGE_HEIGHT_PT * editor.pageZoom}px" style:--page-scale={editor.pageZoom}>
    {@render children()}
  </div>
</div>

<style>
  .page-area { display: grid; place-items: center; min-width: 0; min-height: 0; overflow: auto; background: var(--ok-canvas-bg); }
  .page { position: relative; flex: none; margin: 22px; background: #fff; color: #000; box-shadow: var(--ok-shadow-lg); }
</style>
