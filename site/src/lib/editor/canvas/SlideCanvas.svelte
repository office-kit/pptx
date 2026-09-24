<script lang="ts">
  import DrawingGuides from './DrawingGuides.svelte';
  // The editing surface. Paints the current slide with the preview renderer and
  // manipulates it directly: click/marquee to select, drag to move (multi-shape,
  // with smart-guide snapping), handles to resize, a top handle to rotate,
  // double-click to edit text. Gestures mutate the real model on every frame
  // (so the shape moves for real, not a ghost) via `applyLive`, then commit a
  // single undo step on release. Zoom + right-click menu round out the feel.
  import { onMount, tick } from 'svelte';
  import { tableSelectionBlock, tableCellsInRange } from '../core/table-selection.ts';
  import { parseTableClipboard, canPasteTableCells, pasteTableCells, tableHasMergedCells } from '../core/table-clipboard.ts';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { toggleTextFormat, type TextFormatToggle } from '../core/text-format-toggle.ts';
  import { textEditDiff } from '../core/text-edit-diff.ts';
  import { applyTextFormat, readTextFormat } from '../core/format-clipboard.ts';
  import { parseHtmlTextClipboard, textClipboardHtml } from '../core/html-text-clipboard.ts';
  import { copyTextRange, parseTextClipboard, TEXT_CLIPBOARD_TYPE } from '../core/text-clipboard.ts';
  import { projectTextEdits, replayTextEdits, type TextEdit } from '../core/text-edit-preview.ts';
  import { resolveTextBodyRect, shapeAutoFitScale, shapeCustomTextRect, textColumnsStyle, verticalTextStyle } from '@office-kit/pptx-preview';
  import { shapeTextDefaults } from '../core/text-layout-defaults.ts';
  import { inlineTextHtml } from '../core/inline-text-html.ts';
  import { paragraphsInTextRange } from '../core/paragraph-selection.ts';
  import RichTextInput from '../ui/RichTextInput.svelte';
  import TextFormatBar from '../ui/TextFormatBar.svelte';
  import { t } from '../i18n/i18n.svelte.ts';
  import { getEditor } from '../core/context.ts';
  import {
    getParagraphPropertiesEffective,
    setParagraphAlignment,
    setParagraphLevel,
    setParagraphLineSpacing,
    setParagraphSpacing,
    setParagraphBullet,
    getTableCells,
    getTableCellMargins,
    getTableCellAnchor,
    getShapeBodyPrEffective,
    getShapeBounds,
    getShapeCustomGeometry,
    getShapePreset,
    getShapeTextColumns,
    getShapeTextDirection,
    insertTableRow,
    getTableCellText,
    getTableCellParagraphs,
    setTableCellTextFormat,
    isTableShape,
    getShapeText,
    getShapeFlip,
    getGroupChildren,
    getShapeId,
    getShapeKind,
    getShapeParagraphCount,
    getShapeParagraphElements,
    setShapeTextFormat,
    type TextFormat,
    setShapeBounds,
    setShapeRotation,
  } from '@office-kit/pptx';
  import { selectedShapeIds, selectedShapeId, type Selection } from '../core/selection.ts';
  import { shapeScope, invert, project } from './group-space.ts';
  import { tableCellBoxes, shapeBoxes, slideMetrics, type Box } from './geometry.ts';
  import { resizeRect, resizeSelectionRects, type ResizeHandle } from './resize.ts';
  import { rotateRect, selectionBounds } from './rotation.ts';
  import { type Guide, type Rect } from './snapping.ts';
  import { getGridSpacing, getSnapToGrid } from '@office-kit/pptx';
  import { snapTransformedGrid, snapTransformedMove } from './transformed-snapping.ts';

  const editor = getEditor();
  const doc = editor.doc;

  let stageEl = $state<HTMLDivElement>();
  let areaEl = $state<HTMLDivElement>();

  const metrics = $derived.by(() => {
    doc.version;
    return slideMetrics(doc.pres);
  });

  const gridSpacing = $derived.by(() => { doc.version; return getGridSpacing(doc.pres) ?? { x: 72000, y: 72000 }; });
  const drawingGuides = $derived(editor.guidesVisible() ? editor.drawingGuides() : []);

  // Slide size in CSS px at 96dpi (1 inch = 914400 EMU = 96px).
  const slidePx = $derived.by(() => ({
    w: (metrics.widthEmu / 914400) * 96,
    h: (metrics.heightEmu / 914400) * 96,
  }));

  const stageW = $derived(slidePx.w * editor.zoom);
  const stageH = $derived(slidePx.h * editor.zoom);

  const scope = $derived.by(() => {
    doc.version;
    return doc.currentSlide ? shapeScope(doc.currentSlide, selectedShapeId(doc.selection)) : null;
  });
  const inverseScope = $derived(scope ? invert(scope.matrix) : null);
  const scopeStyle = $derived.by(() => {
    if (!scope) return '';
    const [a, b, c, d, e, f] = scope.matrix;
    return `transform:matrix(${a},${b * pxPerEmuY() / pxPerEmuX()},${c * pxPerEmuX() / pxPerEmuY()},${d},${e * pxPerEmuX()},${f * pxPerEmuY()});transform-origin:0 0;`;
  });
  function localPoint(client: { x: number; y: number }) {
    const rect = stageEl!.getBoundingClientRect();
    const point = {
      x: (client.x - rect.left) / pxPerEmuX(),
      y: (client.y - rect.top) / pxPerEmuY(),
    };
    return inverseScope ? project(inverseScope, point) : point;
  }
  function exitGroup() {
    if (!scope?.parent) return;
    commitEditing();
    doc.selectShape(doc.selection.slideIndex, getShapeId(scope.parent));
  }

  const boxes = $derived.by<Box[]>(() => {
    doc.version;
    const slide = doc.currentSlide;
    if (!slide) return [];
    return shapeBoxes(doc.pres, slide, scope?.shapes ?? []);
  });

  const boxesById = $derived(new Map(boxes.map(box => [box.id, box])));

  const selectedIds = $derived.by<Set<number>>(() => {
    return new Set(selectedShapeIds(doc.selection));
  });

  // Compute a fit-to-area zoom and adopt it until the user zooms themselves.
  function recomputeFit() {
    if (!areaEl) return;
    const avail = areaEl.clientWidth - 56;
    const availH = areaEl.clientHeight - 56;
    const fit = Math.min(avail / slidePx.w, availH / slidePx.h);
    editor.fitZoom = fit > 0 ? fit : 1;
    if (editor.autoFitZoom) editor.zoom = editor.fitZoom;
  }
  $effect(() => {
    slidePx.w;
    recomputeFit();
    if (!areaEl) return;
    const ro = new ResizeObserver(() => recomputeFit());
    ro.observe(areaEl);
    return () => ro.disconnect();
  });
  // ---- Coordinate helpers ------------------------------------------------
  function pxPerEmuX() {
    return stageW / metrics.widthEmu;
  }
  function pxPerEmuY() {
    return stageH / metrics.heightEmu;
  }
  function resolvedRect(id: number): Rect | null {
    const b = boxesById.get(id);
    if (!b) return null;
    return {
      x: (b.left / 100) * metrics.widthEmu,
      y: (b.top / 100) * metrics.heightEmu,
      w: (b.width / 100) * metrics.widthEmu,
      h: (b.height / 100) * metrics.heightEmu,
    };
  }

  // ---- Interaction state -------------------------------------------------
  type Handle = ResizeHandle;
  interface Drag {
    mode: 'move' | 'resize' | 'rotate';
    handle?: Handle;
    ids: number[];
    startClient: { x: number; y: number };
    startRects: Map<number, Rect>;
    startRot: number;
    startRotations: Map<number, number>;
    selectionRect: Rect;
    startAngle: number;
    turn: number;
    center: { x: number; y: number }; // rotate center, EMU
    last: { x: number; y: number };
    shift: boolean;
    moved: boolean;
  }
  let drag = $state<Drag | null>(null);
  const multiFrame = $derived.by(() => {
    if (selectedIds.size < 2) return null;
    if (drag?.mode === 'rotate' && drag.ids.length > 1) return { ...drag.selectionRect, rotation: drag.turn };
    const rects = boxes.filter(box => selectedIds.has(box.id)).map(box => ({ ...resolvedRect(box.id)!, rotation: box.rotation }));
    return rects.length > 1 ? { ...selectionBounds(rects), rotation: 0 } : null;
  });
  let guides = $state<readonly Guide[]>([]);
  let textInput = $state<RichTextInput>();
  let textRange = $state({ start: 0, end: 0 });
  let editing = $state<{ id: number; cell?: { row: number; col: number }; text: string; changes: TextEdit[]; typing?: { format: TextFormat; reset: boolean } } | null>(null);

  // Marquee (rubber-band) selection, in stage-local px.
  let marquee = $state<{ x0: number; y0: number; x1: number; y1: number } | null>(null);

  let cellDrag: { id: number; pointer: number } | null = null;
  const selectedCellBoxes = $derived.by(() => {
    doc.version;
    const selection = doc.selection;
    if (selection.kind !== 'cell') return [];
    const box = boxes.find(box => box.id === selection.shapeId);
    if (!box) return [];
    const cells = getTableCells(box.shape);
    const selected = tableCellsInRange(cells, tableSelectionBlock(selection));
    const flip = getShapeFlip(box.shape);
    return tableCellBoxes(box.shape).filter(cell => selected.has(cells[cell.row]![cell.col]!)).map(cell => ({
      ...cell,
      left: flip?.horizontal ? 100 - cell.left - cell.width : cell.left,
      top: flip?.vertical ? 100 - cell.top - cell.height : cell.top,
    }));
  });

  let gestureSelection: Selection | null = null;
  let cancelling = false;
  function cancelGesture() {
    if (!drag && !marquee && !cellDrag) return;
    if (raf) { cancelAnimationFrame(raf); raf = 0; }
    drag = null;
    marquee = null;
    cellDrag = null;
    guides = [];
    if (gestureSelection) doc.select(gestureSelection);
    gestureSelection = null;
    cancelling = true;
    void doc.cancelLive().catch(cause => editor.toast('error', cause instanceof Error ? cause.message : String(cause))).finally(() => cancelling = false);
  }
  onMount(() => {
    const key = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || (!drag && !marquee && !cellDrag)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      cancelGesture();
    };
    window.addEventListener('keydown', key, true);
    window.addEventListener('blur', cancelGesture);
    return () => {
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('blur', cancelGesture);
      if (raf) cancelAnimationFrame(raf);
    };
  });

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

  // ---- Selection + gesture start ----------------------------------------
  function beginMove(e: PointerEvent, box: Box) {
    if (e.button !== 0 || cancelling) return;
    const selection = doc.selection;
    gestureSelection = selection;
    if (selection.kind === 'cell' && selection.shapeId === box.id) {
      const cell = cellAtPointer(e, box);
      if (cell) {
        e.preventDefault();
        e.stopPropagation();
        commitEditing();
        doc.selectCell(selection.slideIndex, box.id, cell.row, cell.col, e.shiftKey);
        cellDrag = { id: box.id, pointer: e.pointerId };
        if (e.currentTarget instanceof HTMLElement) e.currentTarget.focus({ preventScroll: true });
        capture(e);
        return;
      }
    }
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
    if (e.button !== 0 || cancelling || !inverseScope) return;
    gestureSelection = doc.selection;
    const startRects = new Map<number, Rect>();
    for (const id of ids) {
      const r = resolvedRect(id);
      if (r) startRects.set(id, r);
    }
    const startRotations = new Map(boxes.map(box => [box.id, box.rotation]));
    const selectionRect = ids.length > 1
      ? selectionBounds([...startRects].map(([id, rect]) => ({ ...rect, rotation: startRotations.get(id)! })))
      : startRects.get(ids[0]!)!;
    const center = { x: selectionRect.x + selectionRect.w / 2, y: selectionRect.y + selectionRect.h / 2 };
    const pointer = localPoint({ x: e.clientX, y: e.clientY });
    const startAngle = Math.atan2(pointer.y - center.y, pointer.x - center.x) * 180 / Math.PI;
    drag = {
      mode,
      handle,
      ids,
      startClient: { x: e.clientX, y: e.clientY },
      startRects,
      startRot: startRotations.get(ids[0]!) ?? 0,
      startRotations,
      selectionRect,
      startAngle,
      turn: 0,
      center,
      last: { x: e.clientX, y: e.clientY },
      shift: e.shiftKey,
      moved: false,
    };
    capture(e);
  }

  function onStagePointerDown(e: PointerEvent) {
    if (editing || e.button !== 0 || cancelling) return;
    gestureSelection = doc.selection;
    // Empty-area press → marquee select.
    if (!stageEl) return;
    const pointer = localPoint({ x: e.clientX, y: e.clientY });
    const x = pointer.x * pxPerEmuX();
    const y = pointer.y * pxPerEmuY();
    marquee = { x0: x, y0: y, x1: x, y1: y };
    if (!e.shiftKey && !scope?.parent) doc.clearShapeSelection();
    capture(e);
  }

  function onPointerMove(e: PointerEvent) {
    if (cellDrag) {
      if (cellDrag.pointer !== e.pointerId) return;
      const box = boxes.find(box => box.id === cellDrag!.id);
      const cell = box ? cellAtPointer(e, box) : undefined;
      if (cell) doc.selectCell(doc.selection.slideIndex, cellDrag.id, cell.row, cell.col, true);
      return;
    }
    if (marquee) {
      const pointer = localPoint({ x: e.clientX, y: e.clientY });
      marquee = { ...marquee, x1: pointer.x * pxPerEmuX(), y1: pointer.y * pxPerEmuY() };
      return;
    }
    if (!drag) return;
    if (Math.abs(e.clientX - drag.startClient.x) > 2 || Math.abs(e.clientY - drag.startClient.y) > 2) {
      drag.moved = true;
    }
    drag.last = { x: e.clientX, y: e.clientY };
    drag.shift = e.shiftKey;
    if (drag.moved) schedule(applyDragFrame);
  }

  function applyDragFrame() {
    if (!drag) return;
    const start = localPoint(drag.startClient), last = localPoint(drag.last);
    const dxEmu = last.x - start.x;
    const dyEmu = last.y - start.y;

    if (drag.mode === 'move') {
      const movingIds = new Set(drag.ids);
      const moving = [...drag.startRects].map(([id, rect]) => ({ ...rect, rotation: drag!.startRotations.get(id)! }));
      const others = boxes.filter(box => !movingIds.has(box.id)).flatMap(box => {
        const rect = resolvedRect(box.id);
        return rect ? [{ ...rect, rotation: box.rotation }] : [];
      });
      const delta = getSnapToGrid(doc.pres)
        ? snapTransformedGrid(moving, scope!.matrix, { x: dxEmu, y: dyEmu }, gridSpacing)
        : { x: dxEmu, y: dyEmu };
      const snap = (editor.view.smart || drawingGuides.length > 0) && !getSnapToGrid(doc.pres)
        ? snapTransformedMove(moving, others, scope!.matrix, delta, { w: metrics.widthEmu, h: metrics.heightEmu }, 6 / pxPerEmuX(), { smart: editor.view.smart, drawingGuides })
        : { delta, guides: [] };
      const gdx = snap.delta.x;
      const gdy = snap.delta.y;
      guides = snap.guides;
      doc.applyLive(() => {
        for (const [id, r] of drag!.startRects) {
          const s = doc.shapeById(doc.selection.slideIndex, id);
          if (s) setShapeBounds(s, { x: Math.round(r.x + gdx) as never, y: Math.round(r.y + gdy) as never, w: r.w as never, h: r.h as never });
        }
      });
    } else if (drag.mode === 'resize') {
      const entries = [...drag.startRects];
      const delta = { x: dxEmu, y: dyEmu };
      const minimum = { w: metrics.widthEmu * 0.01, h: metrics.heightEmu * 0.01 };
      const resized = drag.ids.length > 1
        ? resizeSelectionRects(entries.map(([, rect]) => rect), drag.selectionRect, drag.handle!, delta, minimum)
        : [resizeRect(entries[0]![1], drag.handle!, delta, drag.startRot, minimum, drag.shift)];
      guides = [];
      doc.applyLive(() => {
        entries.forEach(([id], index) => {
          const s = doc.shapeById(doc.selection.slideIndex, id);
          const { x, y, w, h } = resized[index]!;
          if (s) setShapeBounds(s, { x: Math.round(x) as never, y: Math.round(y) as never, w: Math.round(w) as never, h: Math.round(h) as never });
        });
      });
    } else {
      const angle = Math.atan2(last.y - drag.center.y, last.x - drag.center.x) * 180 / Math.PI;
      const delta = ((angle - drag.startAngle + 540) % 360) - 180;
      const step = drag.shift ? 15 : 1;
      const base = drag.ids.length === 1 ? drag.startRot : 0;
      drag.turn = Math.round((base + delta) / step) * step - base;
      doc.applyLive(() => {
        for (const [id, bounds] of drag!.startRects) {
          const s = doc.shapeById(doc.selection.slideIndex, id);
          if (!s) continue;
          const rotated = rotateRect({ ...bounds, rotation: drag!.startRotations.get(id)! }, drag!.center, drag!.turn);
          if (drag!.ids.length > 1) setShapeBounds(s, { x: Math.round(rotated.x) as never, y: Math.round(rotated.y) as never, w: bounds.w as never, h: bounds.h as never });
          setShapeRotation(s, rotated.rotation);
        }
      });
    }
  }

  function onPointerUp() {
    gestureSelection = null;
    if (cellDrag) { cellDrag = null; return; }
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
    if (e.button !== 0 || cancelling) return;
    e.stopPropagation();
    doc.selectShape(doc.selection.slideIndex, box.id);
    startDrag('resize', handle, [box.id], e);
  }
  function onRotateDown(e: PointerEvent, box: Box) {
    if (e.button !== 0 || cancelling) return;
    e.stopPropagation();
    doc.selectShape(doc.selection.slideIndex, box.id);
    startDrag('rotate', undefined, [box.id], e);
  }

  function onMultiResizeDown(e: PointerEvent, handle: Handle) {
    if (e.button !== 0 || cancelling) return;
    e.stopPropagation();
    startDrag('resize', handle, [...selectedIds], e);
  }

  function onMultiRotateDown(e: PointerEvent) {
    if (e.button !== 0 || cancelling) return;
    e.stopPropagation();
    startDrag('rotate', undefined, [...selectedIds], e);
  }

  // ---- Text editing ------------------------------------------------------
  function startEditing(box: Box, cell?: { row: number; col: number }) {
    if (getShapeKind(box.shape) === 'picture') {
      doc.selectShape(doc.selection.slideIndex, box.id);
      editor.runOrPrompt('setShapeImageCrop');
      return;
    }
    if (isTableShape(box.shape)) {
      const position = cell ?? (doc.selection.kind === 'cell' ? doc.selection : { row: 0, col: 0 });
      const target = getTableCells(box.shape)[position.row]?.[position.col];
      if (!target) return;
      const text = getTableCellText(target);
      doc.selectCell(doc.selection.slideIndex, box.id, position.row, position.col);
      textRange = { start: text.length, end: text.length };
      editing = { id: box.id, cell: { row: position.row, col: position.col }, text, changes: [] };
      return;
    }
    if (getShapeKind(box.shape) !== 'shape') return;
    let text = '';
    try {
      text = getShapeText(box.shape);
    } catch {
      text = '';
    }
    textRange = { start: text.length, end: text.length };
    editing = { id: box.id, text, changes: [] };
  }

  // Google-Slides parity: with a single shape selected, Enter/F2 edits its text,
  // and simply typing a character enters edit mode replacing the text with it.
  function onTypeToEdit(e: KeyboardEvent) {
    if (editing || e.isComposing || editor.activeDialog) return;
    if (selectedIds.size !== 1) return;
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
      startEditing(box);
      updateEditing(e.key);
    }
  }
  type EditCheckpoint = { text: string; changes: TextEdit[]; changeCount: number; range: { start: number; end: number } };
  let editingUndo: EditCheckpoint[] = [];
  let editingRedo: EditCheckpoint[] = [];
  let historyOwner: typeof editing;
  // Pending text redo belongs to the document state where typing began.
  let editingHistoryDepth = 0;
  let restoringEditing = $state(false);
  let editingHistoryQueue = Promise.resolve();
  let composingText = false;
  let compositionRecorded = false;
  function checkpointEditing() {
    if (!editing || (composingText && compositionRecorded)) return;
    if (composingText) compositionRecorded = true;
    if (historyOwner !== editing || !editing.changes.length) editingUndo = [];
    historyOwner = editing;
    editingUndo.push({ text: editing.text, changes: editing.changes, changeCount: editing.changes.length, range: { ...textRange } });
    editingRedo = [];
    editingHistoryDepth = 0;
  }
  function editingHistory(backward: boolean) {
    const owner = editing;
    editingHistoryQueue = editingHistoryQueue.then(async () => {
      if (!owner || editing !== owner) return;
      await restoreEditingHistory(backward);
    }).catch(cause => editor.toast('error', cause instanceof Error ? cause.message : String(cause)));
  }
  async function restoreEditingHistory(backward: boolean) {
    const cur = editing;
    if (!cur) return;
    if (historyOwner !== cur) { editingUndo = []; editingRedo = []; editingHistoryDepth = 0; historyOwner = cur; }
    const source = backward ? editingUndo : editingRedo;
    const destination = backward ? editingRedo : editingUndo;
    const next = editingHistoryDepth === 0 ? source.pop() : undefined;
    if (next) {
      destination.push({ text: cur.text, changes: cur.changes, changeCount: cur.changes.length, range: { ...textRange } });
      cur.text = next.text;
      cur.changes = next.changes.slice(0, next.changeCount);
      delete cur.typing;
      textRange = next.range;
    } else {
      if (cur.changes.length || !(backward ? doc.canUndo : doc.canRedo)) return;
      const slideIndex = doc.selection.slideIndex;
      const range = { ...textRange };
      const pres = doc.pres;
      restoringEditing = true;
      try {
        await (backward ? doc.undo() : doc.redo());
        if (editing !== cur || doc.pres === pres) return;
        const shape = doc.shapeById(slideIndex, cur.id);
        if (!shape || doc.selection.slideIndex !== slideIndex) { editing = null; return; }
        const cell = cur.cell && isTableShape(shape) ? getTableCells(shape)[cur.cell.row]?.[cur.cell.col] : undefined;
        if (cur.cell && !cell) { editing = null; return; }
        cur.text = cell ? getTableCellText(cell) : getShapeText(shape);
        cur.changes = [];
        delete cur.typing;
        editingHistoryDepth = Math.max(0, editingHistoryDepth + (backward ? 1 : -1));
        textRange = { start: Math.min(range.start, cur.text.length), end: Math.min(range.end, cur.text.length) };
        if (cur.cell) doc.selectCell(slideIndex, cur.id, cur.cell.row, cur.cell.col);
        else doc.selectShape(slideIndex, cur.id);
      } finally {
        restoringEditing = false;
      }
    }
    await tick();
    if (editing === cur) textInput?.setSelectionRange(textRange.start, textRange.end);
  }
  function updateEditing(value: string) {
    if (!editing) return;
    const caretAfter = textInput?.getSelection().start ?? value.length;
    const change = textEditDiff(editing.text, value, textRange, caretAfter);
    if (change) checkpointEditing();
    if (change) editing.changes.push({ ...change, ...(editing.typing ? { typing: { format: { ...editing.typing.format }, reset: editing.typing.reset } } : {}) });
    editing.text = value;
    textRange = { start: caretAfter, end: textInput?.getSelection().end ?? caretAfter };
  }
  function commitEditing() {
    if (!editing) return;
    const box = boxes.find((b) => b.id === editing!.id);
    const cur = editing;
    editing = null;
    if (!box) return;
    if (!cur.changes.length) return;
    doc.transact(t('Edit text'), () => {
      replayEdits(box, cur);
    });
  }

  async function navigateCell(backward: boolean) {
    const cur = editing;
    if (!cur?.cell) return;
    const box = boxes.find(b => b.id === cur.id);
    if (!box) return;
    const cells = tableCellBoxes(box.shape);
    const index = cells.findIndex(cell => cell.row === cur.cell!.row && cell.col === cur.cell!.col);
    let next: { row: number; col: number } | undefined = cells[index + (backward ? -1 : 1)];
    if (!next && backward) return;
    if (!next && tableHasMergedCells(box.shape)) {
      editor.toast('error', t('Split merged cells before adding a row'));
      return;
    }
    if (cur.changes.length || !next) doc.transact(t('Edit table'), () => {
      replayEdits(box, cur);
      if (!next) {
        const row = getTableCells(box.shape).length;
        insertTableRow(box.shape);
        next = { row, col: 0 };
      }
    });
    if (next) startEditing(box, next);
    await tick();
    textInput?.focus();
    textInput?.select();
  }

  function replaceSelectedText(text: string, formats?: TextEdit['formats']) {
    if (!editing) return;
    const start = textInput?.getSelection().start ?? editing.text.length;
    const end = textInput?.getSelection().end ?? start;
    checkpointEditing();
    editing.changes.push({ start, end, text, ...(formats ? { formats } : editing.typing ? { typing: { format: { ...editing.typing.format }, reset: editing.typing.reset } } : {}) });
    if (formats) delete editing.typing;
    editing.text = editing.text.slice(0, start) + text + editing.text.slice(end);
    textRange = { start: start + text.length, end: start + text.length };
    const current = editing;
    void tick().then(() => { if (editing === current) textInput?.setSelectionRange(start + text.length, start + text.length); });
  }
  function copyEditingText(event: ClipboardEvent, cut = false) {
    if (!editing || !pendingTextShape || !textInput || !event.clipboardData) return;
    const { start, end } = textInput.getSelection();
    if (start === end) return;
    const copied = copyTextRange(pendingTextShape, start, end, editing.cell);
    event.clipboardData.setData('text/plain', copied.text);
    event.clipboardData.setData('text/html', textClipboardHtml(copied));
    event.clipboardData.setData(TEXT_CLIPBOARD_TYPE, JSON.stringify(copied));
    event.preventDefault();
    event.stopPropagation();
    if (cut) replaceSelectedText('');
  }
  async function pasteWithoutFormatting() {
    const current = editing;
    const start = textInput?.getSelection().start;
    const end = textInput?.getSelection().end;
    let text: string;
    try { text = await navigator.clipboard.readText(); }
    catch { editor.toast('error', t('Clipboard access was denied')); return; }
    if (editing === current && textInput?.getSelection().start === start && textInput?.getSelection().end === end) replaceSelectedText(text);
  }
  function pasteEditingText(event: ClipboardEvent) {
    if (!editing || !event.clipboardData) return;
    const copied = parseTextClipboard(event.clipboardData.getData(TEXT_CLIPBOARD_TYPE), event.clipboardData.getData('text/plain'));
    if (copied) {
      event.preventDefault();
      event.stopPropagation();
      replaceSelectedText(copied.text, copied.formats);
      return;
    }
    const text = event.clipboardData.getData('text/plain');
    if (!(editing.cell && text.includes('\t'))) {
      const html = parseHtmlTextClipboard(event.clipboardData.getData('text/html'), text);
      if (html) {
        event.preventDefault();
        event.stopPropagation();
        replaceSelectedText(html.text, html.formats);
        return;
      }
    }
    pasteCells(event);
    if (!event.defaultPrevented) {
      event.preventDefault();
      replaceSelectedText(text);
    }
  }

  function pasteCells(event: ClipboardEvent) {
    const cur = editing;
    if (!cur?.cell || !event.clipboardData) return;
    const text = event.clipboardData.getData('text/plain');
    const html = event.clipboardData.getData('text/html');
    if (!text.includes('\t') && !/<table[\s>]/i.test(html)) return;
    const values = parseTableClipboard(text);
    if (!values) return;
    if (values.length === 1 && values[0]!.length === 1) {
      const value = values[0]![0]!;
      if (value === text) return;
      event.preventDefault();
      replaceSelectedText(value);
      return;
    }
    event.preventDefault();
    const box = boxes.find(b => b.id === cur.id);
    if (!box) return;
    if (!canPasteTableCells(box.shape, cur.cell.row, cur.cell.col, values)) {
      editor.toast('error', t('Split merged cells before pasting multiple cells'));
      return;
    }
    doc.transact(t('Paste table cells'), () => {
      pasteTableCells(box.shape, cur.cell!.row, cur.cell!.col, values);
    });
    // The spreadsheet replaces the starting cell, including its pending text.
    startEditing(box, cur.cell);
  }

  function replayEdits(box: Box, cur: NonNullable<typeof editing>) {
    replayTextEdits(box.shape, cur.changes, cur.cell);
  }

  function cellAtPointer(event: MouseEvent, box: Box) {
    if (!isTableShape(box.shape) || !stageEl) return undefined;
    const pointer = localPoint({ x: event.clientX, y: event.clientY });
    const stage = { width: stageW, height: stageH };
    const dx = pointer.x * pxPerEmuX() - (box.left + box.width / 2) / 100 * stage.width;
    const dy = pointer.y * pxPerEmuY() - (box.top + box.height / 2) / 100 * stage.height;
    const angle = box.rotation * Math.PI / 180;
    const flip = getShapeFlip(box.shape);
    const x = (flip?.horizontal ? -1 : 1) * (dx * Math.cos(angle) + dy * Math.sin(angle)) / (box.width / 100 * stage.width) * 100 + 50;
    const y = (flip?.vertical ? -1 : 1) * (-dx * Math.sin(angle) + dy * Math.cos(angle)) / (box.height / 100 * stage.height) * 100 + 50;
    return tableCellBoxes(box.shape).find(c => x >= c.left && x <= c.left + c.width && y >= c.top && y <= c.top + c.height);
  }

  function editAtPointer(event: MouseEvent, box: Box) {
    if (getShapeKind(box.shape) === 'group') {
      const first = getGroupChildren(box.shape)[0];
      if (first) doc.selectShape(doc.selection.slideIndex, getShapeId(first));
      return;
    }
    if (!isTableShape(box.shape)) { startEditing(box); return; }
    const cell = cellAtPointer(event, box);
    if (cell) startEditing(box, cell);
  }

  const editBox = $derived.by(() => {
    const box = boxes.find(b => b.id === editing?.id);
    if (!box || !editing?.cell) return box;
    const cell = tableCellBoxes(box.shape).find(c => c.row === editing?.cell?.row && c.col === editing?.cell?.col);
    if (!cell) return undefined;
    const flip = getShapeFlip(box.shape);
    const dx = (flip?.horizontal ? -1 : 1) * ((cell.left + cell.width / 2) / 100 - 0.5) * box.width / 100 * stageW;
    const dy = (flip?.vertical ? -1 : 1) * ((cell.top + cell.height / 2) / 100 - 0.5) * box.height / 100 * stageH;
    const angle = box.rotation * Math.PI / 180;
    const width = box.width * cell.width / 100;
    const height = box.height * cell.height / 100;
    return { ...box, width, height, left: box.left + box.width / 2 + (dx * Math.cos(angle) - dy * Math.sin(angle)) / stageW * 100 - width / 2, top: box.top + box.height / 2 + (dx * Math.sin(angle) + dy * Math.cos(angle)) / stageH * 100 - height / 2 };
  });

  const textInputStyle = $derived.by(() => {
    const box = editBox;
    if (!box || !scope) return '';
    const base = `left:${box.left}%; top:${box.top}%; width:${box.width}%; height:${box.height}%; transform:rotate(${box.rotation + textBodyTurn}deg);`;

    const [a, b, c, d] = scope.matrix;
    const reflected = a * d - b * c < 0;
    const rotation = box.rotation + textBodyTurn + (getShapeFlip(box.shape)?.vertical ? 180 : 0);
    // Table glyphs follow ancestor scaling; only their reflection is cancelled.
    if (editing?.cell) return `left:${box.left}%; top:${box.top}%; width:${box.width}%; height:${box.height}%; transform:rotate(${rotation}deg) scale(${reflected ? -1 : 1},1); transform-origin:center;`;
    const { x: sx, y: sy } = scope.textScale;
    if (!sx || !sy) return base;
    // Match the preview's text layout: expand the layout box, cancel ancestor
    // scale on glyphs, and cancel reflection before the shape's text rotation.
    return `left:${box.left + box.width * (1 - sx) / 2}%; top:${box.top + box.height * (1 - sy) / 2}%; width:${box.width * sx}%; height:${box.height * sy}%; transform:rotate(${rotation}deg) scale(${(reflected ? -1 : 1) / sx},${1 / sy}); transform-origin:center;`;
  });

  const textBodyStyle = $derived.by(() => {
    doc.version;
    const shape = editBox?.shape;
    if (!shape) return '';
    const cell = editing?.cell;
    const target = cell ? getTableCells(shape)[cell.row]![cell.col]! : null;
    const body = target ? null : getShapeBodyPrEffective(doc.pres, shape);
    const margins = target ? getTableCellMargins(target) : body!.margins;
    const anchor = target
      ? getTableCellAnchor(target) ?? 'top'
      : body!.anchor ?? shapeTextDefaults(shape).anchor;
    let insets = { top: margins.top ?? 45720, right: margins.right ?? 91440, bottom: margins.bottom ?? 45720, left: margins.left ?? 91440 };
    if (!target && editBox && scope) {
      const w = editBox.width / 100 * metrics.widthEmu * scope.textScale.x;
      const h = editBox.height / 100 * metrics.heightEmu * scope.textScale.y;
      // Same rect the renderer lays text into, custom geometry included, so
      // the caret sits where the glyphs will. A custom `<a:rect>` is in the
      // shape's own `<a:ext>` space, so that — not the on-screen box, which
      // carries the group and autofit scales — is what makes it a fraction.
      const rect = resolveTextBodyRect(getShapePreset(shape), { x: 0, y: 0, w, h }, insets, shapeCustomTextRect(getShapeCustomGeometry(shape), getShapeBounds(shape)));
      insets = { left: rect.x, top: rect.y, right: w - rect.x - rect.w, bottom: h - rect.y - rect.h };
    }
    const padding = [insets.top, insets.right, insets.bottom, insets.left]
      .map(value => `${value / 9525 * editor.zoom}px`).join(' ');
    // Vertical writing and multi-column bodies are the renderer's own CSS, so
    // the caret follows the same reading direction as the painted glyphs. The
    // half turn `vert270` needs travels with the box transform instead, which
    // already carries the shape's rotation.
    const vertical = target ? '' : verticalTextStyle(body!.vert ?? getShapeTextDirection(shape)).declarations;
    const columns = target ? '' : textColumnsStyle(getShapeTextColumns(shape));
    // Block alignment keeps literal paragraph separators and selection offsets intact.
    return `padding:${padding}; align-content:${anchor === 'top' ? 'start' : anchor === 'bottom' ? 'safe end' : 'safe center'};${vertical ? ` ${vertical};` : ''}${columns ? ` ${columns};` : ''}`;
  });

  // The half turn that `<a:bodyPr vert="vert270"/>` reads bottom-to-top with.
  const textBodyTurn = $derived.by(() => {
    doc.version;
    const shape = editBox?.shape;
    if (!shape || editing?.cell) return 0;
    const body = getShapeBodyPrEffective(doc.pres, shape);
    return verticalTextStyle(body.vert ?? getShapeTextDirection(shape)).transform ? 180 : 0;
  });

  // The preview shrinks `<a:normAutofit/>` text to fit its box; editing has to
  // shrink by the same factor or the glyphs change size under the caret. The
  // committed model is what the preview painted, so the factor holds for a
  // typing burst and is recomputed when the edit commits.
  const editAutoFit = $derived.by(() => {
    doc.version;
    const box = editBox;
    if (!box || !scope || editing?.cell) return 1;
    return shapeAutoFitScale(doc.pres, box.shape, {
      bounds: {
        x: 0,
        y: 0,
        w: box.width / 100 * metrics.widthEmu * scope.textScale.x,
        h: box.height / 100 * metrics.heightEmu * scope.textScale.y,
      },
    });
  });

  const pendingTextShape = $derived.by(() => {
    doc.version;
    const box = boxes.find(b => b.id === editing?.id);
    return box && editing ? projectTextEdits(box.shape, editing.changes, editing.cell, doc.pres) : null;
  });
  const pendingTextHtml = $derived.by(() => {
    // Formatting mutates OOXML in place, so shape identity alone cannot invalidate this.
    doc.version;
    const shape = pendingTextShape;
    const active = editing;
    if (!shape || !active) return '';
    const source = boxes.find(b => b.id === active.id)?.shape;
    return inlineTextHtml(doc.pres, shape, source, active.cell);
  });
  function selectedTextFormats(shape = boxes.find(b => b.id === editing?.id)?.shape) {
    return shape ? textFormatsInRange(shape, textRange, editing?.cell, { pres: doc.pres, source: boxes.find(b => b.id === editing?.id)?.shape ?? shape }) : [];
  }
  const rangeFormats = $derived.by(() => {
    doc.version;
    const formats = pendingTextShape ? selectedTextFormats(pendingTextShape) : [];
    if (editing?.typing && textRange.start === textRange.end) return [{ ...(editing.typing.reset ? {} : formats[0]), ...editing.typing.format }];
    return formats;
  });
  function inlineParagraphTarget(shape = boxes.find(b => b.id === editing?.id)?.shape) {
    if (!shape || !editing) return null;
    const cell = editing.cell ? getTableCells(shape)[editing.cell.row]![editing.cell.col]! : null;
    const paragraphs = cell ? getTableCellParagraphs(cell).map(p => p.elements)
      : Array.from({ length: getShapeParagraphCount(shape) }, (_, i) => getShapeParagraphElements(shape, i));
    const lengths = paragraphs.map(elements => elements.reduce((length, element) => length + (element.kind === 'br' ? 1 : element.text.length), 0));
    return { shape: cell ?? shape, indices: paragraphsInTextRange(lengths, textRange) };
  }
  const inlineParagraph = $derived.by(() => {
    doc.version;
    const target = pendingTextShape ? inlineParagraphTarget(pendingTextShape) : null;
    const properties = target?.indices.map(index => getParagraphPropertiesEffective(doc.pres, target.shape, index)) ?? [];
    const defaultAlign = pendingTextShape && !editing?.cell ? shapeTextDefaults(pendingTextShape).align : 'left';
    function common(read: (p: typeof properties[number]) => string) {
      const values = properties.map(read);
      return values.every(value => value === values[0]) ? values[0] ?? '' : '';
    }
    return {
      align: common(p => p.align ?? defaultAlign),
      bullet: common(p => typeof p.bullet === 'string' ? p.bullet : p.bullet === null ? 'none' : ''),
      level: common(p => String(p.level)),
      lineKind: common(p => p.lineSpacing?.kind ?? 'inherit'),
      lineValue: common(p => p.lineSpacing ? String(p.lineSpacing.value) : ''),
      before: common(p => p.spcBefPts === null ? '' : String(p.spcBefPts)),
      after: common(p => p.spcAftPts === null ? '' : String(p.spcAftPts)),
    };
  });
  function applyInlineParagraph(kind: 'align' | 'bullet' | 'level' | 'levelDelta' | 'lineKind' | 'lineValue' | 'before' | 'after', value: string) {
    const cur = editing;
    const box = boxes.find(b => b.id === cur?.id);
    if (!cur || !box || restoringEditing) return;
    const range = { ...textRange };
    const lineKind = inlineParagraph.lineKind;
    doc.transact(t('Format paragraphs'), () => {
      replayEdits(box, cur);
      const target = inlineParagraphTarget();
      if (!target) return;
      for (const index of target.indices) {
        if (kind === 'align' && (value === 'left' || value === 'center' || value === 'right' || value === 'justify')) setParagraphAlignment(target.shape, index, value);
        if (kind === 'lineKind' && value === 'inherit') setParagraphLineSpacing(target.shape, index, null);
        if (kind === 'lineKind' && (value === 'pct' || value === 'pts')) setParagraphLineSpacing(target.shape, index, { kind: value, value: value === 'pct' ? 1 : 18 });
        if (kind === 'lineValue' && value !== '' && Number.isFinite(Number(value)) && Number(value) >= 0 && (lineKind === 'pct' || lineKind === 'pts')) setParagraphLineSpacing(target.shape, index, { kind: lineKind, value: Number(value) });
        if ((kind === 'before' || kind === 'after') && (value === '' || (Number.isFinite(Number(value)) && Number(value) >= 0))) setParagraphSpacing(target.shape, index, { [kind === 'before' ? 'beforePts' : 'afterPts']: value === '' ? null : Number(value) });
        if (kind === 'levelDelta') setParagraphLevel(target.shape, index, Math.max(0, Math.min(8, getParagraphPropertiesEffective(doc.pres, target.shape, index).level + Number(value))));
        if (kind === 'level' && /^[0-8]$/.test(value)) setParagraphLevel(target.shape, index, Number(value));
        if (kind === 'bullet' && (value === 'none' || value === 'bullet' || value === 'number')) setParagraphBullet(target.shape, index, value);
      }
    });
    cur.changes = [];
    editingUndo = []; editingRedo = []; editingHistoryDepth = 0;
    void tick().then(() => textInput?.setSelectionRange(range.start, range.end));
  }
  function changeInlineListLevel(delta: number, listsOnly: boolean): boolean {
    const target = pendingTextShape ? inlineParagraphTarget(pendingTextShape) : null;
    if (!target || !target.indices.length) return false;
    const props = target.indices.map(index => getParagraphPropertiesEffective(doc.pres, target.shape, index));
    // Plain-text Tab remains focus navigation; table Tab moves between cells.
    if (listsOnly && props.some(p => p.bullet === null || p.bullet === 'none')) return false;
    if (props.some(p => delta > 0 ? p.level < 8 : p.level > 0)) applyInlineParagraph('levelDelta', String(delta));
    return true;
  }
  function applyInlineFormat(format: TextFormat | ((formats: TextFormat[]) => TextFormat), reset = false) {
    if (!editing || restoringEditing) return;
    if (textRange.start === textRange.end) {
      const resolved = typeof format === 'function' ? format(rangeFormats) : format;
      editing.typing = { format: { ...(reset ? {} : editing.typing?.format), ...resolved }, reset: reset || editing.typing?.reset || false };
      return;
    }
    const cur = editing;
    const box = boxes.find((b) => b.id === cur.id);
    if (!box) return;
    const range = { ...textRange };
    doc.transact(t(reset ? 'Clear text formatting' : 'Format selected text'), () => {
      replayEdits(box, cur);
      const resolved = typeof format === 'function' ? format(selectedTextFormats()) : format;
      if (cur.cell) setTableCellTextFormat(getTableCells(box.shape)[cur.cell.row]![cur.cell.col]!, resolved, { range, reset });
      else setShapeTextFormat(box.shape, resolved, { range, reset });
    });
    cur.changes = [];
    editingUndo = []; editingRedo = []; editingHistoryDepth = 0;
    void tick().then(() => {
      if (editing === cur) textInput?.setSelectionRange(range.start, range.end);
    });
  }
  function toggleInlineFormat(property: TextFormatToggle) {
    applyInlineFormat(formats => toggleTextFormat(formats, property));
  }
  /** Picks up the format at the caret or selection, paragraph included. */
  function copyInlineFormat() {
    const target = pendingTextShape ? inlineParagraphTarget(pendingTextShape) : null;
    if (!target) return;
    const character = (textRange.start === textRange.end ? rangeFormats : selectedTextFormats())[0] ?? {};
    editor.formatClipboard = readTextFormat(doc.pres, target.shape, target.indices[0] ?? 0, character);
    editor.toast('info', t('Formatting copied'));
  }
  /** Repaints the selected range, and every paragraph it touches. */
  function pasteInlineFormat() {
    const format = editor.formatClipboard;
    const cur = editing;
    const box = boxes.find(b => b.id === cur?.id);
    if (!format || !cur || !box || restoringEditing) return;
    const range = { ...textRange };
    doc.transact(t('Paste formatting'), () => {
      replayEdits(box, cur);
      const target = inlineParagraphTarget();
      if (target) applyTextFormat(target.shape, range, target.indices, format, cur.cell !== undefined);
    });
    cur.changes = [];
    editingUndo = []; editingRedo = []; editingHistoryDepth = 0;
    void tick().then(() => { if (editing === cur) textInput?.setSelectionRange(range.start, range.end); });
  }
  function editSelectedTextLink() {
    if (!editing || textRange.start === textRange.end) return;
    const range = { ...textRange };
    const id = editing.id;
    const cell = editing.cell ?? null;
    commitEditing();
    doc.selectShape(doc.selection.slideIndex, id);
    editor.linkTextRange = range;
    editor.linkTableCell = cell;
    editor.runOrPrompt('setShapeHyperlink');
  }
  function onTextFocusOut(event: FocusEvent) {
    const target = event.relatedTarget;
    if (target instanceof Element && target.closest('.canvas-shell .text-format-bar, .inline-edit')) return;
    if (target === null) {
      // Some focus transfers briefly report no related target; inspect the settled focus.
      const current = editing;
      queueMicrotask(() => {
        if (editing === current && !document.activeElement?.closest('.canvas-shell .text-format-bar, .inline-edit')) commitEditing();
      });
      return;
    }
    commitEditing();
  }

  function onContext(e: MouseEvent, box?: Box) {
    // Keep the browser's native text-selection menu while typing.
    if (e.target instanceof Element && e.target.closest('textarea, input, [contenteditable="true"]')) return;
    e.preventDefault();
    e.stopPropagation();
    commitEditing();
    const selection = doc.selection;
    if (!box) doc.clearShapeSelection();
    else if (!(selection.kind === 'shape' && selection.shapeIds.length > 1 && selection.shapeIds.includes(box.id))) {
      const cell = cellAtPointer(e, box);
      if (cell) {
        const cells = getTableCells(box.shape);
        const keepRange = selection.kind === 'cell' && selection.shapeId === box.id &&
          tableCellsInRange(cells, tableSelectionBlock(selection)).has(cells[cell.row]![cell.col]!);
        if (!keepRange) doc.selectCell(selection.slideIndex, box.id, cell.row, cell.col);
      } else if (selection.kind !== 'shape' || !selection.shapeIds.includes(box.id)) {
        doc.selectShape(selection.slideIndex, box.id);
      }
    }
    if (e.currentTarget instanceof HTMLElement) e.currentTarget.focus({ preventScroll: true });
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

<div class="canvas-shell" onfocusout={onTextFocusOut}>
{#if scope?.parent}
  <div class="group-navigation">
    <span>{t('Editing group')}</span>
    <button class="ok-btn" onclick={exitGroup}>{t('Exit group')}</button>
  </div>
{/if}
{#if editing}
  <TextFormatBar formats={rangeFormats} typing selected={textRange.start !== textRange.end} onformat={applyInlineFormat} ontoggle={toggleInlineFormat} paragraph={inlineParagraph} onparagraph={applyInlineParagraph} onlink={editSelectedTextLink} oncopyformat={copyInlineFormat} onpasteformat={pasteInlineFormat} canPasteFormat={!!editor.formatClipboard} ondone={commitEditing} />
{/if}
<div class="canvas-area" bind:this={areaEl} role="presentation">
  <div
    class="stage-wrap"
    style="width:{stageW}px; height:{stageH}px;"
    role="presentation"
    onpointermove={onPointerMove}
    onpointerup={onPointerUp}
    onpointercancel={cancelGesture}
    onlostpointercapture={cancelGesture}
  >
    <div
      bind:this={stageEl}
      class="stage"
      style="width:{stageW}px; height:{stageH}px;"
      onpointerdown={onStagePointerDown}
      oncontextmenu={e => onContext(e)}
      role="presentation"
    >
      {#key doc.selection.slideIndex}
        <div class="paint">{@html doc.currentSvg}</div>
      {/key}

      <div class="overlay">
        {#if editor.view.grid}
          <div class="grid-dots" style="background-size:{(gridSpacing.x * pxPerEmuX()) * Math.max(1, Math.ceil(2 / (gridSpacing.x * pxPerEmuX())))}px {(gridSpacing.y * pxPerEmuY()) * Math.max(1, Math.ceil(2 / (gridSpacing.y * pxPerEmuY())))}px"></div>
        {/if}
        <DrawingGuides guides={drawingGuides} scaleX={pxPerEmuX()} scaleY={pxPerEmuY()} />
        {#each guides as g, i (i)}
          <div class="guide {g.o}" style={guideStyle(g)}></div>
        {/each}
      </div>

      <div class="overlay" style={scopeStyle}>
        {#each boxes as box (box.id)}
          {@const isSel = selectedIds.has(box.id)}
          <div
            class="hit"
            class:selected={isSel}
            style="left:{box.left}%; top:{box.top}%; width:{box.width}%; height:{box.height}%; transform: rotate({box.rotation}deg);"
            role="button"
            tabindex="-1"
            onpointerdown={(e) => beginMove(e, box)}
            oncontextmenu={e => onContext(e, box)}
            ondblclick={(e) => {
              e.stopPropagation();
              editAtPointer(e, box);
            }}
          >
            {#if isSel && doc.selection.kind === 'cell'}
              {#each selectedCellBoxes as cell}
                <div class="cell-selection" aria-hidden="true" style="left:{cell.left}%; top:{cell.top}%; width:{cell.width}%; height:{cell.height}%;"></div>
              {/each}
            {/if}
            {#if isSel && !editing && selectedIds.size === 1}
              <button class="rotate" aria-label={t('Rotate')} title={t('Hold Shift to rotate in 15° steps')} onpointerdown={(e) => onRotateDown(e, box)}></button>
              {#each HANDLES as hd (hd.h)}
                <button
                  class="handle"
                  aria-label={t(`Resize ${hd.h}`)}
                  title={t('Hold Shift to preserve aspect ratio')}
                  style="left:{hd.cx}%; top:{hd.cy}%; cursor:{hd.cur};"
                  onpointerdown={(e) => onHandleDown(e, box, hd.h)}
                ></button>
              {/each}
            {/if}
          </div>
        {/each}

        {#if multiFrame && !editing}
          <div class="multi-selection" style="left:{multiFrame.x * pxPerEmuX()}px; top:{multiFrame.y * pxPerEmuY()}px; width:{multiFrame.w * pxPerEmuX()}px; height:{multiFrame.h * pxPerEmuY()}px; transform: rotate({multiFrame.rotation}deg);">
            <button class="rotate" aria-label={t('Rotate selected objects')} title={t('Hold Shift to rotate in 15° steps')} onpointerdown={onMultiRotateDown}></button>
            {#each HANDLES.filter(handle => handle.h.length === 2) as hd (hd.h)}
              <button class="handle" aria-label={t(`Scale selection ${hd.h}`)} title={t('Resize selection proportionally')}
                style="left:{hd.cx}%; top:{hd.cy}%; cursor:{hd.cur};" onpointerdown={e => onMultiResizeDown(e, hd.h)}></button>
            {/each}
          </div>
        {/if}


        <!-- marquee -->
        {#if marquee}
          <div
            class="marquee"
            style="left:{Math.min(marquee.x0, marquee.x1)}px; top:{Math.min(marquee.y0, marquee.y1)}px; width:{Math.abs(marquee.x1 - marquee.x0)}px; height:{Math.abs(marquee.y1 - marquee.y0)}px;"
          ></div>
        {/if}

        {#if editing}
          {@const eb = editBox}
          {#if eb}
            <RichTextInput
              label={t(editing.cell ? 'Cell text' : 'Edit text')}
              bind:this={textInput}
              onselect={(range) => {
                if (editing && (range.start !== textRange.start || range.end !== textRange.end)) delete editing.typing;
                textRange = range;
              }}
              style={`${textInputStyle} ${textBodyStyle}`}
              value={editing.text}
              html={pendingTextHtml}
              textZoom={editor.zoom * editAutoFit}
              busy={restoringEditing}
              onbeforeinput={(range) => { textRange = range; }}
              oninput={updateEditing}
              onhistory={editingHistory}
              oncomposition={(active) => { composingText = active; compositionRecorded = false; }}
              onnewline={() => replaceSelectedText('\n')}
              oncopy={(event) => copyEditingText(event)}
              oncut={(event) => copyEditingText(event, true)}
              onpaste={pasteEditingText}
              onkeydown={(e) => {
                if (!e.isComposing && (e.ctrlKey || e.metaKey) && e.shiftKey && (e.key.toLowerCase() === 'v' || e.code === 'KeyV')) {
                  e.preventDefault();
                  e.stopPropagation();
                  void pasteWithoutFormatting();
                  return;
                }
                if (!e.isComposing && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
                  commitEditing();
                  return;
                }
                e.stopPropagation();
                if (e.isComposing) return;
                const formatKey = e.key.toLowerCase();
                // Ctrl/Cmd+Alt+C / V, as PowerPoint and Google Slides paint
                // formatting. `code` because Alt rewrites `key` on macOS.
                if ((e.ctrlKey || e.metaKey) && e.altKey && (e.code === 'KeyC' || e.code === 'KeyV')) {
                  e.preventDefault();
                  if (e.code === 'KeyC') copyInlineFormat(); else pasteInlineFormat();
                  return;
                }
                if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.key === '\\') { e.preventDefault(); applyInlineFormat({}, true); }
                if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (formatKey === 'b' || formatKey === 'i' || formatKey === 'u')) { e.preventDefault(); toggleInlineFormat(formatKey === 'b' ? 'bold' : formatKey === 'i' ? 'italic' : 'underline'); }
                else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); editSelectedTextLink(); }
                else if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (e.code === 'BracketLeft' || e.code === 'BracketRight')) { e.preventDefault(); changeInlineListLevel(e.code === 'BracketRight' ? 1 : -1, false); }
                else if (e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey && !editing?.cell && changeInlineListLevel(e.shiftKey ? -1 : 1, true)) e.preventDefault();
                else if (e.key === 'Tab' && editing?.cell) { e.preventDefault(); void navigateCell(e.shiftKey); }
                else if (e.key === 'Escape') editing = null;
                else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) commitEditing();
              }}
            />
          {/if}
        {/if}
      </div>
    </div>
  </div>
</div>

</div>

<style>
  .grid-dots { position: absolute; inset: 0; background-image: radial-gradient(circle, #808080 0.7px, transparent 0.8px); }
  .group-navigation { display: flex; align-items: center; gap: 12px; padding: 4px 12px; background: var(--ok-panel); }
  .canvas-shell { display: flex; flex-direction: column; min-height: 0; min-width: 0; }
  .canvas-area {
    flex: 1;
    background: var(--ok-canvas-bg);
    overflow: auto;
    display: grid;
    place-items: center;
    padding: 28px;
    min-height: 0;
    min-width: 0;
    overflow: auto;
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
  .cell-selection { position: absolute; pointer-events: none; background: color-mix(in srgb, var(--ok-selected-border) 16%, transparent); outline: 2px solid var(--ok-selected-border); outline-offset: -2px; }
  .multi-selection { position: absolute; pointer-events: none; outline: 1px dashed var(--ok-selected-border); transform-origin: center; }
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

</style>
