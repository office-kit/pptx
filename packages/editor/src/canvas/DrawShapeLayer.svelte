<script lang="ts">
  // Draws the shape chosen in the Shapes gallery: drag to size it (Shift keeps
  // it square), or click to drop a one-inch shape, as PowerPoint does.
  import { addSlideLine, emu, getShapeId, inches, type Emu } from '@office-kit/pptx';
  import { addGalleryShape } from '../core/shape-gallery.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { widthEmu, heightEmu }: { widthEmu: number; heightEmu: number } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const CLICK_SIZE = inches(1);
  // A press that moves less than this is a click, not a drag.
  const DRAG_THRESHOLD_PX = 4;

  let layer = $state<HTMLDivElement>();
  let start = $state<{ x: number; y: number; clientX: number; clientY: number } | null>(null);
  let end = $state<{ x: number; y: number } | null>(null);
  let square = false;

  function emuAt(event: PointerEvent) {
    const rect = layer!.getBoundingClientRect();
    return { x: ((event.clientX - rect.left) / rect.width) * widthEmu, y: ((event.clientY - rect.top) / rect.height) * heightEmu };
  }
  function down(event: PointerEvent) {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    layer!.setPointerCapture(event.pointerId);
    start = { ...emuAt(event), clientX: event.clientX, clientY: event.clientY };
    end = null;
  }
  function move(event: PointerEvent) {
    if (!start) return;
    square = event.shiftKey;
    if (Math.hypot(event.clientX - start.clientX, event.clientY - start.clientY) >= DRAG_THRESHOLD_PX) end = emuAt(event);
  }
  const box = $derived.by(() => {
    if (!start || !end) return null;
    let w = end.x - start.x;
    let h = end.y - start.y;
    if (square && editor.drawShape !== 'line') {
      const side = Math.max(Math.abs(w), Math.abs(h));
      w = Math.sign(w || 1) * side;
      h = Math.sign(h || 1) * side;
    }
    return { x1: start.x, y1: start.y, x2: start.x + w, y2: start.y + h };
  });
  function up() {
    const preset = editor.drawShape;
    const slide = doc.currentSlide;
    const from = start;
    const drawn = box;
    start = null;
    end = null;
    if (!preset || !slide || !from) return;
    const area = drawn ?? { x1: from.x, y1: from.y, x2: from.x + CLICK_SIZE, y2: from.y + CLICK_SIZE };
    const point = (value: number) => emu(Math.round(value)) as Emu;
    let id = 0;
    doc.transact(t('Insert Shape'), () => {
      const shape = preset === 'line'
        ? addSlideLine(slide, { from: { x: point(area.x1), y: point(area.y1) }, to: { x: point(area.x2), y: point(area.y2) } })
        : addGalleryShape(slide, preset, {
            x: point(Math.min(area.x1, area.x2)),
            y: point(Math.min(area.y1, area.y2)),
            w: point(Math.max(1, Math.abs(area.x2 - area.x1))),
            h: point(Math.max(1, Math.abs(area.y2 - area.y1))),
          });
      id = getShapeId(shape);
    });
    editor.drawShape = null;
    doc.selectShape(doc.selection.slideIndex, id);
  }
</script>

<div bind:this={layer} class="draw-shape-layer" role="application" aria-label={t('Draw shape')} onpointerdown={down} onpointermove={move} onpointerup={up} onpointercancel={() => { start = null; end = null; }}>
  {#if box}
    <svg viewBox="0 0 {widthEmu} {heightEmu}" preserveAspectRatio="none" aria-hidden="true">
      {#if editor.drawShape === 'line'}
        <line x1={box.x1} y1={box.y1} x2={box.x2} y2={box.y2} style="stroke: var(--ok-accent)" vector-effect="non-scaling-stroke" />
      {:else}
        <rect x={Math.min(box.x1, box.x2)} y={Math.min(box.y1, box.y2)} width={Math.abs(box.x2 - box.x1)} height={Math.abs(box.y2 - box.y1)} fill="rgba(214,51,108,0.08)" style="stroke: var(--ok-accent)" vector-effect="non-scaling-stroke" />
      {/if}
    </svg>
  {/if}
</div>

<style>
  .draw-shape-layer { position: absolute; inset: 0; z-index: 40; pointer-events: auto; cursor: crosshair; touch-action: none; }
  svg { position: absolute; inset: 0; width: 100%; height: 100%; pointer-events: none; }
</style>
