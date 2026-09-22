<script lang="ts">
  // The editing surface. Paints the current slide with the preview renderer and
  // manipulates it directly: click/marquee to select, drag to move (multi-shape,
  // with smart-guide snapping), handles to resize, a top handle to rotate,
  // double-click to edit text. Gestures mutate the real model on every frame
  // (so the shape moves for real, not a ghost) via `applyLive`, then commit a
  // single undo step on release. Zoom + right-click menu round out the feel.
  import { onMount, tick } from 'svelte';
  import { tableSelectionBlock, tableCellsInRange } from '../core/table-selection.ts';
  import { parseTableClipboard, canPasteTableCells, pasteTableCells, tableHasMergedCells } from '../core/table-clipboard.ts';
  import { toggleTextFormat, type TextFormatToggle } from '../core/text-format-toggle.ts';
  import { textEditDiff } from '../core/text-edit-diff.ts';
  import { parseHtmlTextClipboard, textClipboardHtml } from '../core/html-text-clipboard.ts';
  import { copyTextRange, parseTextClipboard, TEXT_CLIPBOARD_TYPE } from '../core/text-clipboard.ts';
  import { projectTextEdits, replayTextEdits, type TextEdit } from '../core/text-edit-preview.ts';
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
    insertTableRow,
    getTableCellText,
    getTableCellParagraphs,
    setTableCellTextFormat,
    isTableShape,
    getShapeText,
    getShapeKind,
    getShapeParagraphCount,
    getShapeParagraphElements,
    setShapeTextFormat,
    type TextFormat,
    setShapeBounds,
    setShapeRotation,
  } from '@office-kit/pptx';
  import { selectedShapeIds, topLevelShapes, type Selection } from '../core/selection.ts';
  import { tableCellBoxes, shapeBoxes, slideMetrics, type Box } from './geometry.ts';
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
    return shapeBoxes(doc.pres, slide, topLevelShapes(slide));
  });

  const selectedIds = $derived.by<Set<number>>(() => {
    return new Set(selectedShapeIds(doc.selection));
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
    return tableCellBoxes(box.shape).filter(cell => selected.has(cells[cell.row]![cell.col]!));
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
    if (e.button !== 0 || cancelling) return;
    gestureSelection = doc.selection;
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
    if (editing || e.button !== 0 || cancelling) return;
    gestureSelection = doc.selection;
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
    if (cellDrag) {
      if (cellDrag.pointer !== e.pointerId) return;
      const box = boxes.find(box => box.id === cellDrag!.id);
      const cell = box ? cellAtPointer(e, box) : undefined;
      if (cell) doc.selectCell(doc.selection.slideIndex, cellDrag.id, cell.row, cell.col, true);
      return;
    }
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
    if (drag.moved) schedule(applyDragFrame);
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
  let composingText = false;
  let compositionRecorded = false;
  function checkpointEditing() {
    if (!editing || (composingText && compositionRecorded)) return;
    if (composingText) compositionRecorded = true;
    if (historyOwner !== editing || !editing.changes.length) editingUndo = [];
    historyOwner = editing;
    editingUndo.push({ text: editing.text, changes: editing.changes, changeCount: editing.changes.length, range: { ...textRange } });
    editingRedo = [];
  }
  function editingHistory(backward: boolean) {
    if (!editing) return;
    if (historyOwner !== editing) { editingUndo = []; editingRedo = []; historyOwner = editing; }
    const source = backward ? editingUndo : editingRedo;
    const destination = backward ? editingRedo : editingUndo;
    const next = source.pop();
    if (!next) return;
    destination.push({ text: editing.text, changes: editing.changes, changeCount: editing.changes.length, range: { ...textRange } });
    editing.text = next.text;
    editing.changes = next.changes.slice(0, next.changeCount);
    delete editing.typing;
    textRange = next.range;
    void tick().then(() => textInput?.setSelectionRange(next.range.start, next.range.end));
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
    const stage = stageEl.getBoundingClientRect();
    const dx = event.clientX - stage.left - (box.left + box.width / 2) / 100 * stage.width;
    const dy = event.clientY - stage.top - (box.top + box.height / 2) / 100 * stage.height;
    const angle = box.rotation * Math.PI / 180;
    const x = (dx * Math.cos(angle) + dy * Math.sin(angle)) / (box.width / 100 * stage.width) * 100 + 50;
    const y = (-dx * Math.sin(angle) + dy * Math.cos(angle)) / (box.height / 100 * stage.height) * 100 + 50;
    return tableCellBoxes(box.shape).find(c => x >= c.left && x <= c.left + c.width && y >= c.top && y <= c.top + c.height);
  }

  function editAtPointer(event: MouseEvent, box: Box) {
    if (!isTableShape(box.shape)) { startEditing(box); return; }
    const cell = cellAtPointer(event, box);
    if (cell) startEditing(box, cell);
  }

  const editBox = $derived.by(() => {
    const box = boxes.find(b => b.id === editing?.id);
    if (!box || !editing?.cell) return box;
    const cell = tableCellBoxes(box.shape).find(c => c.row === editing?.cell?.row && c.col === editing?.cell?.col);
    if (!cell) return undefined;
    const dx = ((cell.left + cell.width / 2) / 100 - 0.5) * box.width / 100 * stageW;
    const dy = ((cell.top + cell.height / 2) / 100 - 0.5) * box.height / 100 * stageH;
    const angle = box.rotation * Math.PI / 180;
    const width = box.width * cell.width / 100;
    const height = box.height * cell.height / 100;
    return { ...box, width, height, left: box.left + box.width / 2 + (dx * Math.cos(angle) - dy * Math.sin(angle)) / stageW * 100 - width / 2, top: box.top + box.height / 2 + (dx * Math.sin(angle) + dy * Math.cos(angle)) / stageH * 100 - height / 2 };
  });

  const pendingTextShape = $derived.by(() => {
    doc.version;
    const box = boxes.find(b => b.id === editing?.id);
    return box && editing ? projectTextEdits(box.shape, editing.changes, editing.cell, doc.pres) : null;
  });
  const pendingTextHtml = $derived.by(() => {
    // Formatting mutates OOXML in place, so shape identity alone cannot invalidate this.
    doc.version;
    return pendingTextShape && editing ? textClipboardHtml(copyTextRange(pendingTextShape, 0, editing.text.length, editing.cell)) : '';
  });
  function selectedTextFormats(shape = boxes.find(b => b.id === editing?.id)?.shape) {
    if (!shape) return [];
    const formats: TextFormat[] = [];
    let offset = 0;
    const paragraphs = editing?.cell
      ? getTableCellParagraphs(getTableCells(shape)[editing.cell.row]![editing.cell.col]!).map(p => p.elements)
      : Array.from({ length: getShapeParagraphCount(shape) }, (_, i) => getShapeParagraphElements(shape, i));
    for (const elements of paragraphs) {
      const paragraphStart = offset;
      for (const element of elements) {
        const length = element.kind === 'br' ? 1 : element.text.length;
        if (textRange.start === textRange.end) {
          const caret = textRange.start;
          if (length && ((offset < caret && offset + length >= caret) || (caret === paragraphStart && offset === caret))) return [element.format ?? {}];
        } else if (offset < textRange.end && offset + length > textRange.start) formats.push(element.format ?? {});
        offset += length;
      }
      offset++;
    }
    return formats;
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
    function common(read: (p: typeof properties[number]) => string) {
      const values = properties.map(read);
      return values.every(value => value === values[0]) ? values[0] ?? '' : '';
    }
    return {
      align: common(p => p.align ?? 'left'),
      bullet: common(p => typeof p.bullet === 'string' ? p.bullet : p.bullet === null ? 'none' : ''),
      level: common(p => String(p.level)),
      lineKind: common(p => p.lineSpacing?.kind ?? 'inherit'),
      lineValue: common(p => p.lineSpacing ? String(p.lineSpacing.value) : ''),
      before: common(p => p.spcBefPts === null ? '' : String(p.spcBefPts)),
      after: common(p => p.spcAftPts === null ? '' : String(p.spcAftPts)),
    };
  });
  function applyInlineParagraph(kind: 'align' | 'bullet' | 'level' | 'lineKind' | 'lineValue' | 'before' | 'after', value: string) {
    const cur = editing;
    const box = boxes.find(b => b.id === cur?.id);
    if (!cur || !box) return;
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
        if (kind === 'level' && /^[0-8]$/.test(value)) setParagraphLevel(target.shape, index, Number(value));
        if (kind === 'bullet' && (value === 'none' || value === 'bullet' || value === 'number')) setParagraphBullet(target.shape, index, value);
      }
    });
    cur.changes = [];
    editingUndo = []; editingRedo = [];
    void tick().then(() => textInput?.setSelectionRange(range.start, range.end));
  }
  function applyInlineFormat(format: TextFormat | ((formats: TextFormat[]) => TextFormat), reset = false) {
    if (!editing) return;
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
    editingUndo = []; editingRedo = [];
    void tick().then(() => {
      if (editing === cur) textInput?.setSelectionRange(range.start, range.end);
    });
  }
  function toggleInlineFormat(property: TextFormatToggle) {
    applyInlineFormat(formats => toggleTextFormat(formats, property));
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
{#if editing}
  <TextFormatBar formats={rangeFormats} typing selected={textRange.start !== textRange.end} onformat={applyInlineFormat} ontoggle={toggleInlineFormat} paragraph={inlineParagraph} onparagraph={applyInlineParagraph} onlink={editSelectedTextLink} ondone={commitEditing} />
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
          {@const eb = editBox}
          {#if eb}
            <RichTextInput
              label={t(editing.cell ? 'Cell text' : 'Edit text')}
              bind:this={textInput}
              onselect={(range) => {
                if (editing && (range.start !== textRange.start || range.end !== textRange.end)) delete editing.typing;
                textRange = range;
              }}
              style="left:{eb.left}%; top:{eb.top}%; width:{eb.width}%; height:{eb.height}%; transform: rotate({eb.rotation}deg);"
              value={editing.text}
              html={pendingTextHtml}
              zoom={editor.zoom}
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
                if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.key === '\\') { e.preventDefault(); applyInlineFormat({}, true); }
                if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (formatKey === 'b' || formatKey === 'i' || formatKey === 'u')) { e.preventDefault(); toggleInlineFormat(formatKey === 'b' ? 'bold' : formatKey === 'i' ? 'italic' : 'underline'); }
                else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); editSelectedTextLink(); }
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
  .canvas-shell { display: flex; flex-direction: column; min-height: 0; min-width: 0; }
  .canvas-area {
    flex: 1;
    background:
      radial-gradient(circle at 1px 1px, rgba(0, 0, 0, 0.05) 1px, transparent 0) 0 0 / 22px 22px,
      var(--ok-canvas-bg);
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
