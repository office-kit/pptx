<script lang="ts">
  // The page of Notes Page view and the handout and notes masters, at the
  // deck's notes page size (7.5 × 10 in portrait by default), fitted to the
  // window with 22 pt around it (the reference desktop app (Mac) opens these views at 96% in a
  // 1512 × 900 pt window), or at the zoom chosen. Contents are placed in page
  // fractions; `--page-scale` is screen px per point.
  import type { Snippet } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { pageSizePt } from '../core/master-geometry.ts';

  let { label, children }: { label: string; children: Snippet } = $props();
  const editor = getEditor();
  const page = $derived.by(() => { editor.doc.version; return pageSizePt(editor.doc.pres); });
  const FIT_MARGIN = 22;
  let areaWidth = $state(0);
  let areaHeight = $state(0);
  $effect(() => {
    const fit = Math.min((areaWidth - 2 * FIT_MARGIN) / page.width, (areaHeight - 2 * FIT_MARGIN) / page.height);
    editor.pageFitZoom = fit > 0 ? fit : 1;
    if (editor.pageAutoFitZoom) editor.pageZoom = editor.pageFitZoom;
  });
</script>

<div class="page-area ok-scroll" role="region" aria-label={label} bind:clientWidth={areaWidth} bind:clientHeight={areaHeight}>
  <div class="page" style:width="{page.width * editor.pageZoom}px" style:height="{page.height * editor.pageZoom}px" style:--page-scale={editor.pageZoom}>
    {@render children()}
  </div>
</div>

<style>
  .page-area { display: grid; place-items: center; min-width: 0; min-height: 0; overflow: auto; background: var(--ok-canvas-bg); }
  .page { position: relative; flex: none; margin: 22px; background: #fff; color: #000; box-shadow: var(--ok-shadow-lg); }
</style>
