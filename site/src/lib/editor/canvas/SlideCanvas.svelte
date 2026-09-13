<script lang="ts">
  // The editing surface. Paints the current slide with the preview renderer and
  // manipulates it directly: click/marquee to select, drag to move (multi-shape,
  // with smart-guide snapping), handles to resize, a top handle to rotate,
  // double-click to edit text. Gestures mutate the real model on every frame
  // (so the shape moves for real, not a ghost) via `applyLive`, then commit a
  // single undo step on release. Zoom + right-click menu round out the feel.
  import { getEditor } from '../core/context.ts';
  import {
    getShapeText,
    getSlideShapes,
    setShapeBounds,
    setShapeRotation,
    setShapeText,
  } from '@office-kit/pptx';
  import { shapeBoxes, slideMetrics, type Box } from './geometry.ts';
  import { snapMove, type Guide, type Rect } from './snapping.ts';

  const editor = getEditor();
  const doc = editor.doc;

  let stageEl = $state<HTMLDivElement>();
  let areaEl = $state<HTMLDivElement>();

  const metrics = $derived.by(() => {
    doc.version;
    return slideMetrics(doc.pres);
  });

  // Slide size in CSS px at 96dpi (1 inch = 914400 EMU = 96px).
  const slidePx = $derived.by(() => ({
    w: (metrics.widthEmu / 914400) * 96,
    h: (metrics.heightEmu / 914400) * 96,
  }));

  const stageW = $derived(slidePx.w * editor.zoom);
  const stageH = $derived(slidePx.h * editor.zoom);

  const boxes = $derived.by<Box[]>(() => {
    doc.version;
    const slide = doc.currentSlide;
    if (!slide) return [];
    return shapeBoxes(doc.pres, slide, getSlideShapes(slide));
  });

  const selectedIds = $derived.by<Set<number>>(() => {
    const sel = doc.selection;
    return sel.kind === 'shape' ? new Set(sel.shapeIds) : new Set<number>();
  });

  // Compute a fit-to-area zoom and adopt it until the user zooms themselves.
  let userZoomed = $state(false);
  function recomputeFit() {
    if (!areaEl) return;
    const avail = areaEl.clientWidth - 56;
    const availH = areaEl.clientHeight - 56;
    const fit = Math.min(avail / slidePx.w, availH / slidePx.h);
    editor.fitZoom = fit > 0 ? fit : 1;
    if (!userZoomed) editor.setZoom(editor.fitZoom);
  }
  $effect(() => {
    slidePx.w;
    recomputeFit();
    if (!areaEl) return;
    const ro = new ResizeObserver(() => recomputeFit());
    ro.observe(areaEl);
    return () => ro.disconnect();
  });
  // Track manual zoom so we stop auto-fitting.
  $effect(() => {
    editor.zoom;
    if (Math.abs(editor.zoom - editor.fitZoom) > 0.001) userZoomed = true;
  });

  // ---- Coordinate helpers ------------------------------------------------
  function pxPerEmuX() {
    return stageW / metrics.widthEmu;
  }
  function pxPerEmuY() {
    return stageH / metrics.heightEmu;
  }
  function resolvedRect(id: number): Rect | null {
    const b = boxes.find((x) => x.id === id);
    if (!b) return null;
    return {
      x: (b.left / 100) * metrics.widthEmu,
      y: (b.top / 100) * metrics.heightEmu,
      w: (b.width / 100) * metrics.widthEmu,
      h: (b.height / 100) * metrics.heightEmu,
    };
  }

  // ---- Interaction state -------------------------------------------------
  type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
  interface Drag {
    mode: 'move' | 'resize' | 'rotate';
    handle?: Handle;
    ids: number[];
    startClient: { x: number; y: number };
    startRects: Map<number, Rect>;
    startRot: number;
    center: { x: number; y: number }; // rotate center, EMU
    last: { x: number; y: number };
    shift: boolean;
    moved: boolean;
  }
  let drag = $state<Drag | null>(null);
  let guides = $state<readonly Guide[]>([]);
  let editing = $state<{ id: number; text: string } | null>(null);

  // Marquee (rubber-band) selection, in stage-local px.
  let marquee = $state<{ x0: number; y0: number; x1: number; y1: number } | null>(null);

  let raf = 0;
  function schedule(fn: () => void) {
    if (raf) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      fn();
    });
  }

  function capture(e: PointerEvent) {
    try {
      (e.target as HTMLElement).setPointerCapture?.(e.pointerId);
    } catch {
      /* ignore */
    }
  }

  // Reliably focus the inline editor when it mounts (the `autofocus` attribute
  // does not fire for dynamically-inserted nodes). Places the caret at the end.
  function focusEdit(node: HTMLTextAreaElement) {
    requestAnimationFrame(() => {
      node.focus();
      const len = node.value.length;
      node.setSelectionRange(len, len);
    });
  }

  // ---- Selection + gesture start ----------------------------------------
  function beginMove(e: PointerEvent, box: Box) {
    if (editing) return;
    e.stopPropagation();
    const already = selectedIds.has(box.id);
    if (e.shiftKey) {
      doc.selectShape(doc.selection.slideIndex, box.id, true);
      return;
    }
    if (!already) doc.selectShape(doc.selection.slideIndex, box.id);
    const ids = selectedIds.has(box.id) && selectedIds.size > 1 ? [...selectedIds] : [box.id];
    startDrag('move', undefined, ids, e);
  }

  function startDrag(mode: Drag['mode'], handle: Handle | undefined, ids: number[], e: PointerEvent) {
    const startRects = new Map<number, Rect>();
    for (const id of ids) {
      const r = resolvedRect(id);
      if (r) startRects.set(id, r);
    }
    const first = startRects.get(ids[0]!);
    drag = {
      mode,
      handle,
      ids,
      startClient: { x: e.clientX, y: e.clientY },
      startRects,
      startRot: boxes.find((b) => b.id === ids[0])?.rotation ?? 0,
      center: first ? { x: first.x + first.w / 2, y: first.y + first.h / 2 } : { x: 0, y: 0 },
      last: { x: e.clientX, y: e.clientY },
      shift: e.shiftKey,
      moved: false,
    };
    capture(e);
  }

  function onStagePointerDown(e: PointerEvent) {
    if (editing || e.button !== 0) return;
    // Empty-area press → marquee select.
    if (!stageEl) return;
    const rect = stageEl.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    marquee = { x0: x, y0: y, x1: x, y1: y };
    if (!e.shiftKey) doc.clearShapeSelection();
    capture(e);
  }

  function onPointerMove(e: PointerEvent) {
    if (marquee) {
      const rect = stageEl!.getBoundingClientRect();
      marquee = { ...marquee, x1: e.clientX - rect.left, y1: e.clientY - rect.top };
      return;
    }
    if (!drag) return;
    if (Math.abs(e.clientX - drag.startClient.x) > 2 || Math.abs(e.clientY - drag.startClient.y) > 2) {
      drag.moved = true;
    }
    drag.last = { x: e.clientX, y: e.clientY };
    drag.shift = e.shiftKey;
    schedule(applyDragFrame);
  }

  function applyDragFrame() {
    if (!drag) return;
    const dxEmu = (drag.last.x - drag.startClient.x) / pxPerEmuX();
    const dyEmu = (drag.last.y - drag.startClient.y) / pxPerEmuY();

    if (drag.mode === 'move') {
      // Snap the group's bounding box, then move every shape by the same delta.
      const rects = [...drag.startRects.values()];
      const bx = Math.min(...rects.map((r) => r.x));
      const by = Math.min(...rects.map((r) => r.y));
      const bw = Math.max(...rects.map((r) => r.x + r.w)) - bx;
      const bh = Math.max(...rects.map((r) => r.y + r.h)) - by;
      const others = boxes.filter((b) => !drag!.ids.includes(b.id)).map((b) => resolvedRect(b.id)!).filter(Boolean);
      const thresh = 6 / pxPerEmuX();
      const snap = snapMove({ x: bx + dxEmu, y: by + dyEmu, w: bw, h: bh }, others, { w: metrics.widthEmu, h: metrics.heightEmu }, thresh);
      const gdx = snap.x - bx;
      const gdy = snap.y - by;
      guides = snap.guides;
      doc.applyLive(() => {
        for (const [id, r] of drag!.startRects) {
          const s = doc.shapeById(doc.selection.slideIndex, id);
          if (s) setShapeBounds(s, { x: Math.round(r.x + gdx) as never, y: Math.round(r.y + gdy) as never, w: r.w as never, h: r.h as never });
        }
      });
    } else if (drag.mode === 'resize') {
      const id = drag.ids[0]!;
      const r = drag.startRects.get(id)!;
      let { x, y, w, h } = r;
      const hd = drag.handle!;
      if (hd.includes('e')) w = r.w + dxEmu;
      if (hd.includes('s')) h = r.h + dyEmu;
      if (hd.includes('w')) { x = r.x + dxEmu; w = r.w - dxEmu; }
      if (hd.includes('n')) { y = r.y + dyEmu; h = r.h - dyEmu; }
      w = Math.max(w, metrics.widthEmu * 0.01);
      h = Math.max(h, metrics.heightEmu * 0.01);
      guides = [];
      doc.applyLive(() => {
        const s = doc.shapeById(doc.selection.slideIndex, id);
        if (s) setShapeBounds(s, { x: Math.round(x) as never, y: Math.round(y) as never, w: Math.round(w) as never, h: Math.round(h) as never });
      });
    } else {
      const id = drag.ids[0]!;
      const rect = stageEl!.getBoundingClientRect();
      const cx = rect.left + (drag.center.x / metrics.widthEmu) * stageW;
      const cy = rect.top + (drag.center.y / metrics.heightEmu) * stageH;
      const ang = (Math.atan2(drag.last.y - cy, drag.last.x - cx) * 180) / Math.PI + 90;
      const step = drag.shift ? 15 : 1;
      const snapped = (Math.round(((ang + 360) % 360) / step) * step) % 360;
      doc.applyLive(() => {
        const s = doc.shapeById(doc.selection.slideIndex, id);
        if (s) setShapeRotation(s, snapped);
      });
    }
  }

  function onPointerUp() {
    if (raf) {
      cancelAnimationFrame(raf);
      raf = 0;
    }
    if (marquee) {
      finishMarquee();
      marquee = null;
      return;
    }
    if (drag) {
      const moved = drag.moved;
      const label = drag.mode === 'move' ? 'Move' : drag.mode === 'resize' ? 'Resize' : 'Rotate';
      if (moved) {
        applyDragFrame();
        doc.commit(label); // a real gesture → one undo step
      }
      drag = null;
      guides = [];
    }
  }

  function finishMarquee() {
    if (!marquee || !stageEl) return;
    const x0 = Math.min(marquee.x0, marquee.x1);
    const y0 = Math.min(marquee.y0, marquee.y1);
    const x1 = Math.max(marquee.x0, marquee.x1);
    const y1 = Math.max(marquee.y0, marquee.y1);
    if (x1 - x0 < 4 && y1 - y0 < 4) return; // a click, not a drag
    // to EMU
    const mx0 = (x0 / stageW) * metrics.widthEmu;
    const my0 = (y0 / stageH) * metrics.heightEmu;
    const mx1 = (x1 / stageW) * metrics.widthEmu;
    const my1 = (y1 / stageH) * metrics.heightEmu;
    const hits = boxes.filter((b) => {
      const bx = (b.left / 100) * metrics.widthEmu;
      const by = (b.top / 100) * metrics.heightEmu;
      const bw = (b.width / 100) * metrics.widthEmu;
      const bh = (b.height / 100) * metrics.heightEmu;
      return bx < mx1 && bx + bw > mx0 && by < my1 && by + bh > my0;
    });
    if (hits.length) {
      doc.select({ kind: 'shape', slideIndex: doc.selection.slideIndex, shapeIds: hits.map((h) => h.id) });
    }
  }

  // ---- Handles / rotate --------------------------------------------------
  function onHandleDown(e: PointerEvent, box: Box, handle: Handle) {
    e.stopPropagation();
    doc.selectShape(doc.selection.slideIndex, box.id);
    startDrag('resize', handle, [box.id], e);
  }
  function onRotateDown(e: PointerEvent, box: Box) {
    e.stopPropagation();
    doc.selectShape(doc.selection.slideIndex, box.id);
    startDrag('rotate', undefined, [box.id], e);
  }

  // ---- Text editing ------------------------------------------------------
  function startEditing(box: Box) {
    let text = '';
    try {
      text = getShapeText(box.shape);
    } catch {
      text = '';
    }
    editing = { id: box.id, text };
  }

  // Google-Slides parity: with a single shape selected, Enter/F2 edits its text,
  // and simply typing a character enters edit mode replacing the text with it.
  function onTypeToEdit(e: KeyboardEvent) {
    if (editing) return;
    if (doc.selection.kind !== 'shape' || selectedIds.size !== 1) return;
    const t = e.target as HTMLElement;
    if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName))) return;
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    const box = boxes.find((b) => selectedIds.has(b.id));
    if (!box) return;
    if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      startEditing(box);
    } else if (e.key.length === 1) {
      e.preventDefault();
      editing = { id: box.id, text: e.key };
    }
  }
  function commitEditing() {
    if (!editing) return;
    const box = boxes.find((b) => b.id === editing!.id);
    const cur = editing;
    editing = null;
    if (!box) return;
    doc.transact('Edit text', () => setShapeText(box.shape, cur.text));
  }

  function onContext(e: MouseEvent) {
    e.preventDefault();
    editor.openContextMenu(e.clientX, e.clientY);
  }

  const HANDLES: { h: Handle; cx: number; cy: number; cur: string }[] = [
    { h: 'nw', cx: 0, cy: 0, cur: 'nwse-resize' },
    { h: 'n', cx: 50, cy: 0, cur: 'ns-resize' },
    { h: 'ne', cx: 100, cy: 0, cur: 'nesw-resize' },
    { h: 'e', cx: 100, cy: 50, cur: 'ew-resize' },
    { h: 'se', cx: 100, cy: 100, cur: 'nwse-resize' },
    { h: 's', cx: 50, cy: 100, cur: 'ns-resize' },
    { h: 'sw', cx: 0, cy: 100, cur: 'nesw-resize' },
    { h: 'w', cx: 0, cy: 50, cur: 'ew-resize' },
  ];

  function guideStyle(g: Guide): string {
    if (g.o === 'v') {
      const left = (g.pos / metrics.widthEmu) * 100;
      const top = (g.from / metrics.heightEmu) * 100;
      const height = ((g.to - g.from) / metrics.heightEmu) * 100;
      return `left:${left}%; top:${top}%; height:${height}%; width:0;`;
    }
    const top = (g.pos / metrics.heightEmu) * 100;
    const left = (g.from / metrics.widthEmu) * 100;
    const width = ((g.to - g.from) / metrics.widthEmu) * 100;
    return `top:${top}%; left:${left}%; width:${width}%; height:0;`;
  }
</script>

<svelte:window onkeydown={onTypeToEdit} />

<div class="canvas-area" bind:this={areaEl} role="presentation">
  <div
    class="stage-wrap"
    style="width:{stageW}px; height:{stageH}px;"
    role="presentation"
    onpointermove={onPointerMove}
    onpointerup={onPointerUp}
  >
    <div
      bind:this={stageEl}
      class="stage"
      style="width:{stageW}px; height:{stageH}px;"
      onpointerdown={onStagePointerDown}
      oncontextmenu={onContext}
      role="presentation"
    >
      {#key doc.selection.slideIndex}
        <div class="paint">{@html doc.currentSvg}</div>
      {/key}

      <div class="overlay">
        {#each boxes as box (box.id)}
          {@const isSel = selectedIds.has(box.id)}
          <div
            class="hit"
            class:selected={isSel}
            style="left:{box.left}%; top:{box.top}%; width:{box.width}%; height:{box.height}%; transform: rotate({box.rotation}deg);"
            role="button"
            tabindex="-1"
            onpointerdown={(e) => beginMove(e, box)}
            ondblclick={(e) => {
              e.stopPropagation();
              startEditing(box);
            }}
          >
            {#if isSel && !editing && selectedIds.size === 1}
              <button class="rotate" aria-label="Rotate" onpointerdown={(e) => onRotateDown(e, box)}></button>
              {#each HANDLES as hd (hd.h)}
                <button
                  class="handle"
                  aria-label={`Resize ${hd.h}`}
                  style="left:{hd.cx}%; top:{hd.cy}%; cursor:{hd.cur};"
                  onpointerdown={(e) => onHandleDown(e, box, hd.h)}
                ></button>
              {/each}
            {/if}
          </div>
        {/each}

        <!-- smart guides -->
        {#each guides as g, i (i)}
          <div class="guide {g.o}" style={guideStyle(g)}></div>
        {/each}

        <!-- marquee -->
        {#if marquee}
          <div
            class="marquee"
            style="left:{Math.min(marquee.x0, marquee.x1)}px; top:{Math.min(marquee.y0, marquee.y1)}px; width:{Math.abs(marquee.x1 - marquee.x0)}px; height:{Math.abs(marquee.y1 - marquee.y0)}px;"
          ></div>
        {/if}

        {#if editing}
          {@const eb = boxes.find((b) => b.id === editing?.id)}
          {#if eb}
            <textarea
              class="inline-edit"
              style="left:{eb.left}%; top:{eb.top}%; width:{eb.width}%; height:{eb.height}%;"
              bind:value={editing.text}
              use:focusEdit
              onpointerdown={(e) => e.stopPropagation()}
              onpointerup={(e) => e.stopPropagation()}
              ondblclick={(e) => e.stopPropagation()}
              onblur={commitEditing}
              onkeydown={(e) => {
                e.stopPropagation();
                if (e.key === 'Escape') editing = null;
                else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) commitEditing();
              }}
            ></textarea>
          {/if}
        {/if}
      </div>
    </div>
  </div>
</div>

<style>
  .canvas-area {
    background:
      radial-gradient(circle at 1px 1px, rgba(0, 0, 0, 0.05) 1px, transparent 0) 0 0 / 22px 22px,
      var(--ok-canvas-bg);
    overflow: auto;
    display: grid;
    place-items: center;
    padding: 28px;
    min-height: 0;
  }
  .stage-wrap {
    position: relative;
    flex: none;
  }
  .stage {
    position: relative;
    background: #fff;
    box-shadow: var(--ok-shadow-lg);
    touch-action: none;
  }
  .paint,
  .overlay {
    position: absolute;
    inset: 0;
  }
  .paint :global(svg) {
    width: 100%;
    height: 100%;
    display: block;
  }
  .overlay {
    pointer-events: none;
  }
  .hit {
    position: absolute;
    pointer-events: auto;
    cursor: move;
    transform-origin: center;
  }
  .hit:hover:not(.selected) {
    outline: 1px solid rgba(43, 108, 176, 0.5);
  }
  .hit.selected {
    outline: 1.5px solid var(--ok-selected-border);
  }
  .handle {
    position: absolute;
    width: 10px;
    height: 10px;
    margin: -5px 0 0 -5px;
    background: #fff;
    border: 1.5px solid var(--ok-selected-border);
    border-radius: 2px;
    padding: 0;
    pointer-events: auto;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.2);
  }
  .rotate {
    position: absolute;
    left: 50%;
    top: -24px;
    width: 13px;
    height: 13px;
    margin-left: -6px;
    border-radius: 50%;
    background: #fff;
    border: 1.5px solid var(--ok-selected-border);
    cursor: grab;
    padding: 0;
    pointer-events: auto;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.2);
  }
  .guide {
    position: absolute;
    background: #e5407a;
    pointer-events: none;
    z-index: 6;
  }
  .guide.v {
    width: 1px !important;
    margin-left: -0.5px;
  }
  .guide.h {
    height: 1px !important;
    margin-top: -0.5px;
  }
  .marquee {
    position: absolute;
    border: 1px solid var(--ok-selected-border);
    background: rgba(43, 108, 176, 0.12);
    pointer-events: none;
  }
  .inline-edit {
    position: absolute;
    pointer-events: auto;
    border: 1px solid var(--ok-selected-border);
    background: rgba(255, 255, 255, 0.96);
    font-family: var(--ok-font);
    font-size: 14px;
    padding: 4px;
    resize: none;
    z-index: 7;
  }
</style>
