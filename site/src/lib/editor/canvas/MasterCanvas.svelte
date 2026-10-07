<script lang="ts">
  // Slide Master view's editing area: the selected master or layout at the
  // slide zoom, fitted to the window like the slide canvas. A layout's
  // placeholders can be dragged to new positions, which every slide on the
  // layout inherits; the rest is edited through the Slide Master tab.
  import { emu, getSlideLayoutName, getSlideLayoutPartName, getSlideSize, type ShapeBounds } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { layoutBoxes, masterGroups } from '../core/master-geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import MasterSlide from '../ui/MasterSlide.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  // As on the slide canvas: 100% is one slide point per screen point, and Fit
  // leaves 22 pt around the slide.
  const EMU_PER_PT = 12700;
  const FIT_MARGIN = 22;
  const size = $derived.by(() => { doc.version; return getSlideSize(doc.pres) ?? { width: 12192000, height: 6858000 }; });
  const target = $derived.by(() => {
    doc.version;
    const partName = doc.layoutTarget?.partName ?? null;
    for (const group of masterGroups(doc.pres)) {
      if (partName === null && group.partName === editor.selectedMaster && group.layouts[0]) return { layout: group.layouts[0], master: true, label: t('Slide Master') };
      const layout = group.layouts.find((item) => getSlideLayoutPartName(item) === partName);
      if (layout) return { layout, master: false, label: getSlideLayoutName(layout) };
    }
    return null;
  });
  const boxes = $derived(target && !target.master ? layoutBoxes(target.layout, size) : []);
  let areaWidth = $state(0);
  let areaHeight = $state(0);
  $effect(() => {
    const fit = Math.min((areaWidth - 2 * FIT_MARGIN) / (size.width / EMU_PER_PT), (areaHeight - 2 * FIT_MARGIN) / (size.height / EMU_PER_PT));
    editor.fitZoom = fit > 0 ? fit : 1;
    if (editor.autoFitZoom) editor.zoom = editor.fitZoom;
  });
  const width = $derived((size.width / EMU_PER_PT) * editor.zoom);
  const height = $derived((size.height / EMU_PER_PT) * editor.zoom);

  /** Moves placeholder `index` by (dx, dy) EMU, writing its layout transform. */
  function place(index: number, dx: number, dy: number) {
    const box = boxes[index]!;
    const bounds: ShapeBounds = { x: emu(box.x * size.width + dx), y: emu(box.y * size.height + dy), w: emu(box.w * size.width), h: emu(box.h * size.height) };
    editor.invoke('setSlideLayoutPlaceholderBounds', { index, bounds });
  }
  let drag = $state<{ index: number; pointer: number; x: number; y: number; dx: number; dy: number } | null>(null);
  function start(event: PointerEvent, index: number) {
    if (event.button !== 0) return;
    event.preventDefault();
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    drag = { index, pointer: event.pointerId, x: event.clientX, y: event.clientY, dx: 0, dy: 0 };
  }
  function move(event: PointerEvent) {
    if (drag?.pointer !== event.pointerId) return;
    drag = { ...drag, dx: event.clientX - drag.x, dy: event.clientY - drag.y };
  }
  function end(event: PointerEvent) {
    if (drag?.pointer !== event.pointerId) return;
    const { index, dx, dy } = drag;
    drag = null;
    if (!boxes[index] || (dx === 0 && dy === 0)) return;
    place(index, (dx / width) * size.width, (dy / height) * size.height);
  }
  function nudge(event: KeyboardEvent, index: number) {
    const step = event.shiftKey ? 182880 : 18288;
    const delta = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[event.key];
    if (!delta || !boxes[index]) return;
    event.preventDefault();
    place(index, delta[0]!, delta[1]!);
  }
</script>

<div class="master-canvas ok-scroll" bind:clientWidth={areaWidth} bind:clientHeight={areaHeight} role="region" aria-label={t('Slide Editor Pane')}>
  {#if target}
    <div class="stage" role="img" aria-label={target.label} style:width="{width}px" style:height="{height}px">
      <MasterSlide pres={doc.pres} layout={target.layout} master={target.master} version={doc.version} pixelWidth={width} />
      {#each boxes as box, index (index)}
        {@const moving = drag?.index === index ? drag : null}
        <button class="handle" aria-label="{t('Placeholder')} {index + 1}" style:left="{box.x * width + (moving?.dx ?? 0)}px" style:top="{box.y * height + (moving?.dy ?? 0)}px" style:width="{box.w * width}px" style:height="{box.h * height}px" onpointerdown={(event) => start(event, index)} onpointermove={move} onpointerup={end} onpointercancel={() => (drag = null)} onkeydown={(event) => nudge(event, index)}></button>
      {/each}
    </div>
  {/if}
</div>

<style>
  .master-canvas { display: grid; place-items: center; min-width: 0; min-height: 0; overflow: auto; background: var(--ok-canvas-bg); }
  .stage { position: relative; flex: none; margin: 22px; background: #fff; box-shadow: var(--ok-shadow-lg); }
  .handle { position: absolute; box-sizing: border-box; padding: 0; border: 0; background: transparent; cursor: move; touch-action: none; }
  .handle:focus-visible, .handle:active { outline: 1px solid var(--ok-selected-border); }
</style>
