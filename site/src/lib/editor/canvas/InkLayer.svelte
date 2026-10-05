<script lang="ts">
  // Draw-tab tools: freehand pen strokes, the stroke eraser and Lasso Select.
  // The layer covers the slide while a tool is on, so presses never reach the
  // shapes underneath (as in PowerPoint, where drawing never moves a shape).
  import { getEditor } from '../core/context.ts';
  import { addInkStroke, addRecognizedShape, inkAt, recognizeShape, shapesInLasso, type Point } from '../core/ink.svelte.ts';
  import { removeShape } from '@office-kit/pptx';
  import { t } from '../i18n/i18n.svelte.ts';

  let { widthEmu, heightEmu }: { widthEmu: number; heightEmu: number } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const ink = editor.ink;

  let layer = $state<HTMLDivElement>();
  let points = $state<Point[]>([]);
  let erasing = false;

  const ERASER_PX = 6;

  function emuAt(event: PointerEvent): Point {
    const rect = layer!.getBoundingClientRect();
    return {
      x: ((event.clientX - rect.left) / rect.width) * widthEmu,
      y: ((event.clientY - rect.top) / rect.height) * heightEmu,
    };
  }
  let lastErase: Point | null = null;
  // Pointer events arrive far apart on a fast swipe; test the path between
  // them too, so the eraser removes every stroke it crossed.
  function erase(event: PointerEvent) {
    const slide = doc.currentSlide;
    if (!slide) return;
    const tolerance = (ERASER_PX / layer!.getBoundingClientRect().width) * widthEmu;
    const to = emuAt(event);
    const from = lastErase ?? to;
    lastErase = to;
    const steps = Math.max(1, Math.ceil(Math.hypot(to.x - from.x, to.y - from.y) / tolerance));
    for (let i = 1; i <= steps; i++) {
      const point = { x: from.x + ((to.x - from.x) * i) / steps, y: from.y + ((to.y - from.y) * i) / steps };
      const stroke = inkAt(slide, point, tolerance);
      if (stroke) doc.applyLive(() => removeShape(stroke));
    }
  }
  function down(event: PointerEvent) {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    layer!.setPointerCapture(event.pointerId);
    if (ink.tool === 'eraser') {
      erasing = true;
      lastErase = null;
      erase(event);
    } else points = [emuAt(event)];
  }
  function move(event: PointerEvent) {
    if (erasing) erase(event);
    else if (points.length) points = [...points, emuAt(event)];
  }
  function up() {
    if (erasing) {
      erasing = false;
      if (doc.liveEditing) doc.commit(t('Erase'));
      return;
    }
    const stroke = points;
    points = [];
    const slide = doc.currentSlide;
    if (!stroke.length || !slide) return;
    if (ink.tool === 'pen') {
      doc.clearShapeSelection();
      const recognized = ink.toShape ? recognizeShape(stroke) : null;
      doc.transact(t('Draw'), () =>
        recognized ? addRecognizedShape(slide, recognized, ink.pen) : addInkStroke(slide, stroke, ink.pen),
      );
    } else {
      const ids = shapesInLasso(slide, stroke);
      doc.clearShapeSelection();
      ids.forEach((id, index) => doc.selectShape(doc.selection.slideIndex, id, index > 0));
      ink.tool = null;
    }
  }
  function cancel() {
    points = [];
    if (erasing) {
      erasing = false;
      if (doc.liveEditing) doc.commit(t('Erase'));
    }
  }
  const pathData = $derived(points.map((p, i) => `${i ? 'L' : 'M'}${p.x} ${p.y}`).join(' '));
</script>

<svelte:window onkeydown={(event) => { if (event.key === 'Escape') { cancel(); ink.tool = null; } }} />

<div
  bind:this={layer}
  class="ink-layer"
  class:eraser={ink.tool === 'eraser'}
  role="application"
  aria-label={t(ink.tool === 'pen' ? 'Draw' : ink.tool === 'eraser' ? 'Eraser' : 'Lasso Select')}
  data-ink-tool={ink.tool}
  onpointerdown={down}
  onpointermove={move}
  onpointerup={up}
  onpointercancel={cancel}
>
  {#if points.length}
    <svg viewBox="0 0 {widthEmu} {heightEmu}" preserveAspectRatio="none" aria-hidden="true">
      {#if ink.tool === 'pen'}
        <path d={pathData} fill="none" stroke={ink.pen.color} stroke-opacity={ink.pen.opacity} stroke-width={ink.pen.widthEmu} stroke-linecap="round" stroke-linejoin="round" />
      {:else}
        <path d="{pathData} Z" fill="rgba(0,120,212,0.08)" stroke="#0078d4" stroke-dasharray="4 3" vector-effect="non-scaling-stroke" />
      {/if}
    </svg>
  {/if}
</div>

<style>
  .ink-layer { position: absolute; inset: 0; z-index: 40; pointer-events: auto; cursor: crosshair; touch-action: none; }
  .ink-layer.eraser { cursor: cell; }
  svg { position: absolute; inset: 0; width: 100%; height: 100%; pointer-events: none; }
</style>
