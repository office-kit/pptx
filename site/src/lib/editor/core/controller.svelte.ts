// The editor controller — the façade the UI talks to.
//
// It owns the document, runs commands through the registry, and holds the
// transient UI state that several components share (which command dialog is
// open, whether the palette is showing, the toast queue). Ribbon buttons,
// context menus and the palette all funnel through `invoke` / `runOrPrompt`,
// so there is exactly one path from "user intent" to "library call".

import { lockedShapeIds, selectionLocked } from './shape-locks.ts';
import {
  setShapeLocked,
  setShapeZIndex,
  groupShapes,
  getDrawingGuides,
  getDrawingGuidesVisible,
  setDrawingGuides,
  type DrawingGuide,
  getTableCells,
  getTableCellText,
  setTableCellText,
  copyShape,
  importSlide,
  moveSlide,
  addSlideImage,
  getShapeKind,
  getShapeRotation,
  setShapeImage,
  emu,
  loadPresentation,
  getSlides,
  getSlideShapes,
  getSlideSize,
  findShapeById,
  type PresentationData,
  getShapeBoundsResolved,
  getShapeId,
  inches,
  removeShape,
  setShapeBounds,
  type SlideShapeData,
  type ShapeBounds,
} from '@office-kit/pptx';
import {
  copyTableCellValues,
  serializeTableClipboard,
  canPasteTableCells,
  pasteTableCells,
} from './table-clipboard.ts';
import { neighboringTableCell, tableCellsInRange, tableSelectionBlock } from './table-selection.ts';
import {
  applyShapeFormat,
  copiedFormatLimits,
  readShapeFormat,
  type CopiedFormat,
} from './format-clipboard.ts';
import { hideSlideNumber, showSlideNumber, slideNumberShape } from './slide-numbers.ts';
import { textFormatsInRange } from './text-format-selection.ts';
import { getCommand, type Command, type CommandContext } from './registry.ts';
import { capabilityById } from '../manifest/index.ts';
import { EditorDocument } from './document.svelte.ts';
import { t } from '../i18n/i18n.svelte.ts';
import { ViewPreferences } from './view-preferences.svelte.ts';
import { shapeScope, invert, project, type Matrix } from '../canvas/group-space.ts';
import { projectedBounds } from '../canvas/transformed-snapping.ts';
import type { Rect } from '../canvas/snapping.ts';
import { selectedSlideIndices, selectedShapeId, selectedShapeIds } from './selection.ts';

export interface Toast {
  readonly id: number;
  readonly kind: 'info' | 'error';
  readonly message: string;
}

export interface ContextMenuState {
  readonly x: number;
  readonly y: number;
}

/** A byte snapshot keeps copied content independent of live edits and history. */
interface Clipboard {
  presentation: Promise<PresentationData>;
  slideIndex: number;
  content: { kind: 'slide'; slideIndices: number[] } | { kind: 'shapes'; shapeIds: number[] };
}

const PASTE_OFFSET = inches(0.25);

let toastSeq = 0;

export class EditorController {
  readonly doc = new EditorDocument();
  readonly view = new ViewPreferences();
  ribbonVisible = $state(true);
  thumbnailsVisible = $state(true);
  selectionPaneVisible = $state(false);
  alignmentReference = $state<'selection' | 'slide'>('selection');
  rotationFocusRequested = $state(false);
  thumbnailWidth = $state<number | null>(null);
  viewMode = $state<'normal' | 'sorter'>('normal');
  sorterZoom = $state(1);
  notesVisible = $state(false);
  notesHeight = $state(120);
  notesFocusRequest = $state(0);

  showNotes(): void {
    this.setViewMode('normal');
    this.notesVisible = true;
    this.notesFocusRequest++;
  }

  showRotationOptions(): void {
    this.setViewMode('normal');
    this.selectionPaneVisible = false;
    this.rotationFocusRequested = true;
  }

  setViewMode(mode: 'normal' | 'sorter'): void {
    this.contextMenu = null;
    this.viewMode = mode;
  }

  drawingGuides(): readonly DrawingGuide[] {
    this.doc.version;
    const size = getSlideSize(this.doc.pres);
    return (
      getDrawingGuides(this.doc.pres) ??
      (size
        ? [
            { id: 1, axis: 'x', position: size.width / 2, color: '#808080' },
            { id: 2, axis: 'y', position: size.height / 2, color: '#808080' },
          ]
        : [])
    );
  }

  guidesVisible(): boolean {
    this.doc.version;
    return this.view.drawing ?? getDrawingGuidesVisible(this.doc.pres) ?? false;
  }

  addDrawingGuide(axis: 'x' | 'y'): void {
    const size = getSlideSize(this.doc.pres);
    if (!size) return;
    const guides = this.drawingGuides();
    const id = Math.max(0, ...guides.map((item) => item.id)) + 1;
    const position =
      (axis === 'x' ? size.width : size.height) / 2 +
      91440 * guides.filter((item) => item.axis === axis).length;
    this.doc.transact(t('Edit guides'), () =>
      setDrawingGuides(this.doc.pres, [...guides, { id, axis, position, color: '#888888' }]),
    );
    this.view.save({ grid: this.view.grid, smart: this.view.smart, drawing: true });
  }

  /** Command whose argument dialog is currently open (null = none). */
  activeDialog = $state<string | null>(null);
  linkTableCell = $state<{ row: number; col: number } | null>(null);
  linkTextRange = $state<{ start: number; end: number } | null>(null);
  paletteOpen = $state<boolean>(false);
  /** The format painter's pickup, shared by the canvas and the panels. */
  formatClipboard = $state<CopiedFormat | null>(null);
  toasts = $state<Toast[]>([]);

  get ctx(): CommandContext {
    return { doc: this.doc };
  }

  command(id: string): Command | undefined {
    return getCommand(id);
  }

  canRun(id: string): boolean {
    this.doc.version;
    const cmd = getCommand(id);
    return cmd ? cmd.canRun(this.ctx) : false;
  }

  reorderMembers(): SlideShapeData[] {
    this.doc.version;
    const slide = this.doc.currentSlide;
    const ids = new Set(selectedShapeIds(this.doc.selection));
    if (!slide || ids.size < 2) return [];
    const siblings = shapeScope(slide, ids.values().next().value ?? null).shapes;
    const members = siblings.filter((shape) => ids.has(getShapeId(shape)));
    return members.length === ids.size ? members : [];
  }

  reorderSelection(frontToBack: readonly number[]): void {
    const members = this.reorderMembers();
    const selected = new Set(members.map(getShapeId));
    if (
      members.length !== frontToBack.length ||
      new Set(frontToBack).size !== selected.size ||
      frontToBack.some((id) => !selected.has(id))
    )
      return;
    const slide = this.doc.currentSlide!;
    const siblings = shapeScope(slide, frontToBack[0] ?? null).shapes;
    const byId = new Map(members.map((shape) => [getShapeId(shape), shape]));
    const backToFront = [...frontToBack].reverse();
    let index = 0;
    const ordered = siblings.map((shape) =>
      selected.has(getShapeId(shape)) ? byId.get(backToFront[index++]!)! : shape,
    );
    if (ordered.every((shape, i) => getShapeId(shape) === getShapeId(siblings[i]!))) return;
    this.doc.transact(t('Reorder Objects'), () => setShapeZIndex(ordered, 0));
  }

  private regroupMembers(): SlideShapeData[] {
    this.doc.version;
    const selection = this.doc.selection;
    const slide = this.doc.currentSlide;
    const members =
      slide && selection.kind === 'shape'
        ? this.doc.regroupHistory.members(slide, selection.shapeIds)
        : [];
    const locked = slide ? lockedShapeIds(slide) : new Set<number>();
    return members.some((shape) => locked.has(getShapeId(shape))) ? [] : members;
  }

  canRegroup(): boolean {
    if (this.selectionLocked()) return false;
    return this.regroupMembers().length >= 2;
  }

  regroupSelection(): void {
    const members = this.regroupMembers();
    if (members.length < 2 || this.selectionLocked()) return;
    this.doc.transact(t('Regroup'), () => {
      const group = groupShapes(members);
      this.doc.regroupHistory.forget(this.doc.currentSlide!, members);
      this.doc.selectShape(this.doc.selection.slideIndex, getShapeId(group));
    });
  }

  /** Execute a command immediately with fully-supplied args. */
  invoke(id: string, args: Record<string, unknown> = {}): unknown {
    const cmd = getCommand(id);
    if (!cmd) {
      this.toast('error', `Unknown command: ${id}`);
      return undefined;
    }
    if (!cmd.canRun(this.ctx)) {
      const cap = capabilityById.get(id);
      this.toast('error', `${cap?.labelEn ?? id}: select a ${cap?.operand ?? 'target'} first.`);
      return undefined;
    }
    try {
      return cmd.run(this.ctx, args);
    } catch (err) {
      this.toast('error', `${id}: ${(err as Error).message}`);
      return undefined;
    }
  }

  /**
   * Run a command, but if it still needs user-supplied required arguments,
   * open its argument dialog instead. `presetArgs` can pre-fill some params
   * (e.g. a color chosen in the ribbon).
   */
  runOrPrompt(id: string, presetArgs: Record<string, unknown> = {}): void {
    if (id === 'setSlideNotes' && presetArgs.value === undefined) {
      this.showNotes();
      return;
    }
    const cmd = getCommand(id);
    if (!cmd) return;
    const needs = cmd.params.filter((p) => !p.optional && presetArgs[p.name] === undefined);
    if (needs.length === 0) {
      this.invoke(id, presetArgs);
    } else {
      this.pendingPreset = presetArgs;
      this.activeDialog = id;
    }
  }

  /** Place decoded image bytes as one undoable edit, or replace one selected picture. */
  applyImage(
    bytes: Uint8Array,
    name: string,
    width: number,
    height: number,
    replace = false,
  ): boolean {
    const slide = this.doc.currentSlide;
    if (!slide) return false;
    if (![width, height].every((value) => Number.isFinite(value) && value > 0)) {
      this.toast('error', t('The image could not be read'));
      return false;
    }
    const selection = this.doc.selection;
    try {
      if (replace) {
        const ids = selectedShapeIds(selection);
        const shape = ids.length === 1 ? findShapeById(slide, ids[0]!) : null;
        if (!shape || getShapeKind(shape) !== 'picture') return false;
        this.doc.transact(t('Replace image'), () => setShapeImage(shape, bytes));
      } else {
        const size = getSlideSize(this.doc.pres);
        if (!size) return false;
        const scale = Math.min((size.width * 0.8) / width, (size.height * 0.8) / height);
        const w = emu(Math.max(1, Math.round(width * scale)));
        const h = emu(Math.max(1, Math.round(height * scale)));
        this.doc.transact(t('Insert image'), () => {
          const shape = addSlideImage(slide, bytes, {
            name,
            x: emu(Math.round((size.width - w) / 2)),
            y: emu(Math.round((size.height - h) / 2)),
            w,
            h,
          });
          this.doc.selectShape(getSlides(this.doc.pres).indexOf(slide), getShapeId(shape));
        });
      }
      return true;
    } catch (error) {
      this.toast(
        'error',
        `${t('The image could not be inserted')}: ${error instanceof Error ? error.message : String(error)}`,
      );
      return false;
    }
  }

  /** Args seeded into a dialog opened via runOrPrompt. */
  pendingPreset: Record<string, unknown> = {};

  closeDialog(): void {
    this.activeDialog = null;
    this.linkTextRange = null;
    this.linkTableCell = null;
    this.pendingPreset = {};
  }

  togglePalette(open?: boolean): void {
    this.paletteOpen = open ?? !this.paletteOpen;
  }

  toast(kind: Toast['kind'], message: string): void {
    const id = ++toastSeq;
    this.toasts = [...this.toasts, { id, kind, message }];
    setTimeout(
      () => {
        this.toasts = this.toasts.filter((t) => t.id !== id);
      },
      kind === 'error' ? 6000 : 3000,
    );
  }

  // --- Zoom --------------------------------------------------------------
  /** Canvas zoom multiplier (1 = fit-ish base). */
  zoom = $state(1);
  autoFitZoom = $state(true);
  /** When set by the canvas, `fit` recomputes to this multiplier. */
  fitZoom = $state(1);

  setZoom(z: number): void {
    if (this.viewMode === 'sorter') this.sorterZoom = Math.max(0.1, Math.min(z, 4));
    else {
      this.autoFitZoom = false;
      this.zoom = Math.max(0.1, Math.min(z, 4));
    }
  }
  zoomIn(): void {
    this.setZoom(
      (Math.round((this.viewMode === 'sorter' ? this.sorterZoom : this.zoom) * 10) + 1) / 10,
    );
  }
  zoomOut(): void {
    this.setZoom(
      (Math.round((this.viewMode === 'sorter' ? this.sorterZoom : this.zoom) * 10) - 1) / 10,
    );
  }
  zoomFit(): void {
    if (this.viewMode === 'sorter') this.sorterZoom = 1;
    else {
      this.autoFitZoom = true;
      this.zoom = this.fitZoom;
    }
  }

  // --- Context menu ------------------------------------------------------
  contextMenu = $state<ContextMenuState | null>(null);
  openContextMenu(x: number, y: number): void {
    this.contextMenu = { x, y };
  }
  closeContextMenu(): void {
    this.contextMenu = null;
  }

  // --- Clipboard & shape actions -----------------------------------------
  #clipboard = $state.raw<Clipboard | { values: string[][] } | null>(null);

  /** Resolve the currently selected shapes to live objects. */
  selectedShapes(): SlideShapeData[] {
    const sel = this.doc.selection;
    const ids = selectedShapeIds(sel);
    const slide = this.doc.slideAt(sel.slideIndex);
    if (!slide || !ids.length) return [];
    const shapes = new Map(getSlideShapes(slide).map((shape) => [getShapeId(shape), shape]));
    return ids.map((id) => shapes.get(id)).filter((s): s is SlideShapeData => s != null);
  }

  moveCellSelection(dr: number, dc: number, extend = false): void {
    const selection = this.doc.selection;
    if (selection.kind !== 'cell') return;
    const table = this.doc.shapeById(selection.slideIndex, selection.shapeId);
    if (!table) return;
    const current = selection.end ?? selection;
    const next = neighboringTableCell(table, current.row, current.col, dr, dc);
    if (next)
      this.doc.selectCell(selection.slideIndex, selection.shapeId, next.row, next.col, extend);
  }

  clearCellText(): void {
    const selection = this.doc.selection;
    if (selection.kind !== 'cell') return;
    const table = this.doc.shapeById(selection.slideIndex, selection.shapeId);
    if (!table) return;
    const cells = [
      ...tableCellsInRange(getTableCells(table), tableSelectionBlock(selection)),
    ].filter((cell) => getTableCellText(cell));
    if (!cells.length) return;
    this.doc.transact(t('Clear cell text'), () => {
      for (const cell of cells) setTableCellText(cell, '', { preserveFormatting: true });
    });
  }

  selectionLocked(): boolean {
    this.doc.version;
    return selectionLocked(this.doc.currentSlide, this.doc.selection);
  }

  lockObjects(shapes: readonly SlideShapeData[], locked: boolean): void {
    if (!shapes.length) return;
    this.doc.transact(t(locked ? 'Lock objects' : 'Unlock objects'), () =>
      setShapeLocked(shapes, locked),
    );
  }

  private selectedGeometry() {
    const items: { shape: SlideShapeData; bounds: ShapeBounds; visible: Rect; inverse: Matrix }[] =
      [];
    const slide = this.doc.currentSlide;
    if (!slide || this.selectionLocked()) return items;
    const scope = shapeScope(slide, selectedShapeId(this.doc.selection));
    const inverse = invert(scope.matrix);
    if (!inverse) return items;
    for (const shape of this.selectedShapes()) {
      const bounds = getShapeBoundsResolved(this.doc.pres, shape);
      if (!bounds) {
        this.toast('error', t('The selection contains an object without a position or size'));
        return [];
      }
      items.push({
        shape,
        bounds,
        visible: projectedBounds({ ...bounds, rotation: getShapeRotation(shape) }, scope.matrix),
        inverse,
      });
    }
    return items;
  }

  /** Align visible edges to the chosen reference; one object always uses the slide. */
  alignSelection(
    alignment: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom',
    reference: 'selection' | 'slide' = 'selection',
  ): void {
    const items = this.selectedGeometry();
    if (!items.length) return;
    const size = getSlideSize(this.doc.pres);
    const toSlide = items.length === 1 || reference === 'slide';
    if (toSlide && !size) {
      this.toast('error', t('Slide size is unavailable'));
      return;
    }
    const boxes = items.map((item) => item.visible);
    const left = toSlide ? 0 : Math.min(...boxes.map((b) => b.x));
    const top = toSlide ? 0 : Math.min(...boxes.map((b) => b.y));
    const right = toSlide ? size!.width : Math.max(...boxes.map((b) => b.x + b.w));
    const bottom = toSlide ? size!.height : Math.max(...boxes.map((b) => b.y + b.h));
    this.doc.transact(t('Align objects'), () => {
      for (const { shape, bounds, visible: b, inverse } of items) {
        let x: number = b.x;
        let y: number = b.y;
        if (alignment === 'left') x = left;
        if (alignment === 'center') x = (left + right - b.w) / 2;
        if (alignment === 'right') x = right - b.w;
        if (alignment === 'top') y = top;
        if (alignment === 'middle') y = (top + bottom - b.h) / 2;
        if (alignment === 'bottom') y = bottom - b.h;
        const origin = project(inverse, { x: 0, y: 0 });
        const delta = project(inverse, { x: x - b.x, y: y - b.y });
        setShapeBounds(shape, {
          ...bounds,
          x: emu(Math.round(bounds.x + delta.x - origin.x)),
          y: emu(Math.round(bounds.y + delta.y - origin.y)),
        });
      }
    });
  }

  /** Mac PowerPoint includes the outer margins when distributing to the slide. */
  distributeSelection(direction: 'horizontal' | 'vertical'): void {
    const items = this.selectedGeometry();
    const toSlide = items.length === 1 || this.alignmentReference === 'slide';
    if (!items.length || (!toSlide && items.length < 3)) return;
    const size = getSlideSize(this.doc.pres);
    if (toSlide && !size) {
      this.toast('error', t('Slide size is unavailable'));
      return;
    }
    const axis = direction === 'horizontal' ? 'x' : 'y';
    const extent = direction === 'horizontal' ? 'w' : 'h';
    items.sort((a, b) => a.visible[axis] - b.visible[axis]);
    const first = items[0]!.visible;
    const last = items[items.length - 1]!.visible;
    const total = items.reduce((sum, item) => sum + item.visible[extent], 0);
    const gap = toSlide
      ? ((axis === 'x' ? size!.width : size!.height) - total) / (items.length + 1)
      : (last[axis] + last[extent] - first[axis] - total) / (items.length - 1);
    this.doc.transact(t('Distribute objects'), () => {
      let position: number = toSlide ? gap : first[axis];
      for (let index = 0; index < items.length; index++) {
        const { shape, bounds, visible: b, inverse } = items[index]!;
        if (toSlide || (index > 0 && index < items.length - 1)) {
          const origin = project(inverse, { x: 0, y: 0 });
          const delta = project(inverse, {
            x: axis === 'x' ? position - b.x : 0,
            y: axis === 'y' ? position - b.y : 0,
          });
          setShapeBounds(shape, {
            ...bounds,
            x: emu(Math.round(bounds.x + delta.x - origin.x)),
            y: emu(Math.round(bounds.y + delta.y - origin.y)),
          });
        }
        position += b[extent] + gap;
      }
    });
  }

  selectAll(): void {
    const selection = this.doc.selection;
    if (selection.kind !== 'cell') {
      this.selectAllShapes();
      return;
    }
    const table = this.doc.shapeById(selection.slideIndex, selection.shapeId);
    if (!table) return;
    const cells = getTableCells(table);
    const columns = cells[0]?.length ?? 0;
    if (!cells.length || !columns) return;
    this.doc.selectCell(selection.slideIndex, selection.shapeId, 0, 0);
    this.doc.selectCell(
      selection.slideIndex,
      selection.shapeId,
      cells.length - 1,
      columns - 1,
      true,
    );
  }

  selectAllShapes(): void {
    const slideIndex = this.doc.selection.slideIndex;
    const slide = this.doc.slideAt(slideIndex);
    if (!slide) return;
    const ids = shapeScope(slide, selectedShapeId(this.doc.selection)).shapes.map((s) =>
      getShapeId(s),
    );
    if (ids.length) this.doc.select({ kind: 'shape', slideIndex, shapeIds: ids });
  }

  /**
   * Copies the formatting of one selected object — paint, and the character
   * and paragraph formatting its text starts with. Text editing has its own
   * path in the canvas, which samples the selected range instead.
   */
  copyObjectFormat(): void {
    const shape = this.selectedShapes()[0];
    if (!shape) return;
    // The first run's effective format stands for the object's text, the way
    // PowerPoint's own format painter treats a whole-object pickup.
    const character =
      textFormatsInRange(shape, { start: 0, end: Number.MAX_SAFE_INTEGER }, undefined, {
        pres: this.doc.pres,
        source: shape,
      })[0] ?? null;
    this.formatClipboard = readShapeFormat(this.doc.pres, shape, character);
    const limits = copiedFormatLimits(this.doc.pres, shape);
    this.toast(
      'info',
      limits.length
        ? `${t('Formatting copied')} — ${t('not copied')}: ${limits.map((limit) => t(limit)).join(', ')}`
        : t('Formatting copied'),
    );
  }

  /** Pastes the copied formatting onto every selected object. */
  pasteObjectFormat(): void {
    const format = this.formatClipboard;
    const shapes = this.selectedShapes();
    if (!format || !shapes.length) return;
    this.doc.transact(t('Paste formatting'), () => {
      for (const shape of shapes) applyShapeFormat(shape, format);
    });
  }

  /** True when every slide in the deck shows a live slide number. */
  slideNumbersOn(): boolean {
    const slides = getSlides(this.doc.pres);
    return slides.length > 0 && slides.every((slide) => slideNumberShape(slide) !== null);
  }

  /**
   * Turns slide numbers on or off for the whole deck, as one undo step —
   * Google Slides' Insert ▸ Slide numbers, which is deck-wide rather than a
   * field the author inserts slide by slide.
   */
  setSlideNumbers(on: boolean): void {
    const slides = getSlides(this.doc.pres);
    if (!slides.length) return;
    let missed = 0;
    this.doc.transact(t(on ? 'Show slide numbers' : 'Hide slide numbers'), () => {
      for (const slide of slides) {
        if (!on) hideSlideNumber(slide);
        else if (showSlideNumber(this.doc.pres, slide) === null) missed += 1;
      }
    });
    if (missed) this.toast('error', t('Slide size is unavailable'));
  }

  deleteSelection(): void {
    if (this.doc.selection.kind === 'cell') {
      this.clearCellText();
      return;
    }
    if (this.doc.selection.kind === 'slide') {
      this.invoke('removeSlide');
      return;
    }
    const shapes = this.selectedShapes();
    if (!shapes.length) return;
    this.doc.transact(t('Delete'), () => {
      for (const s of shapes) removeShape(s);
      this.doc.clearShapeSelection();
    });
  }

  /** Clone shapes onto `slide`, offset, and return the new ids. */
  #cloneOnto(shapes: SlideShapeData[], slideIndex: number, offset: number): number[] {
    const slide = this.doc.slideAt(slideIndex);
    if (!slide) return [];
    const newIds: number[] = [];
    for (const src of shapes) {
      const copy = copyShape(slide, src, { preserveGroupTransform: true });
      const b = getShapeBoundsResolved(this.doc.pres, copy);
      if (b) {
        setShapeBounds(copy, {
          x: emu(b.x + offset),
          y: emu(b.y + offset),
          w: b.w,
          h: b.h,
        });
      }
      newIds.push(getShapeId(copy));
    }
    return newIds;
  }

  duplicateSelection(): void {
    if (this.doc.selection.kind === 'slide') {
      this.invoke('duplicateSlide');
      return;
    }
    const shapes = this.selectedShapes();
    if (!shapes.length) return;
    const slideIndex = this.doc.selection.slideIndex;
    this.doc.transact(t('Duplicate'), () => {
      const newIds = this.#cloneOnto(shapes, slideIndex, PASTE_OFFSET);
      this.doc.select({ kind: 'shape', slideIndex, shapeIds: newIds });
    });
  }

  copySelection(): string | undefined {
    const sel = this.doc.selection;
    if (sel.kind === 'cell') {
      const table = this.doc.shapeById(sel.slideIndex, sel.shapeId);
      if (!table) return;
      const values = copyTableCellValues(table, sel);
      if (!values.length) return;
      this.#clipboard = { values };
      return serializeTableClipboard(values);
    }
    const ids = selectedShapeIds(sel);
    if (sel.kind !== 'slide' && !ids.length) return;
    if (!this.doc.slideAt(sel.slideIndex)) return;
    this.#clipboard = {
      presentation: this.doc.toBytes().then(loadPresentation),
      slideIndex: sel.slideIndex,
      content:
        sel.kind === 'slide'
          ? { kind: 'slide', slideIndices: selectedSlideIndices(sel) }
          : { kind: 'shapes', shapeIds: [...ids] },
    };
    // Attach a handler immediately; a later paste reports the captured failure.
    void this.#clipboard.presentation.catch((error: Error) => this.toast('error', error.message));
  }

  cutSelection(): string | undefined {
    if (this.doc.selection.kind !== 'slide' && !this.selectedShapes().length) return;
    const text = this.copySelection();
    this.deleteSelection();
    return text;
  }

  async paste(): Promise<void> {
    const clip = this.#clipboard;
    if (!clip) return;
    if ('values' in clip) {
      this.pasteCellValues(clip.values);
      return;
    }
    const targetPresentation = this.doc.pres;
    const targetSlide = this.doc.slideAt(this.doc.selection.slideIndex);
    if (!targetSlide && clip.content.kind !== 'slide') return;
    try {
      const source = await clip.presentation;
      // New/Open or undo replaces the document. Never paste into a stale target.
      if (this.doc.pres !== targetPresentation) return;
      const slideIndex = targetSlide ? getSlides(targetPresentation).indexOf(targetSlide) : -1;
      if (targetSlide && slideIndex < 0) return;
      const sourceSlide = getSlides(source)[clip.slideIndex]!;
      if (clip.content.kind === 'slide') {
        this.doc.transact(t('Paste'), () => {
          const inserted: number[] = [];
          for (const sourceIndex of clip.content.kind === 'slide'
            ? clip.content.slideIndices
            : []) {
            const imported = importSlide(targetPresentation, getSlides(source)[sourceIndex]!);
            const at = slideIndex + 1 + inserted.length;
            moveSlide(targetPresentation, imported, at);
            inserted.push(at);
          }
          this.doc.select({
            kind: 'slide',
            slideIndex: inserted[0]!,
            slideIndices: inserted,
            anchorIndex: inserted[0],
          });
        });
        return;
      }
      const sources = clip.content.shapeIds.map((id) => findShapeById(sourceSlide, id)!);
      this.doc.transact(t('Paste'), () => {
        const newIds = this.#cloneOnto(sources, slideIndex, PASTE_OFFSET);
        this.doc.select({ kind: 'shape', slideIndex, shapeIds: newIds });
      });
    } catch (error) {
      this.toast(
        'error',
        `${t('Paste')}: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }

  pasteCellValues(values: string[][]): void {
    const selection = this.doc.selection;
    if (selection.kind !== 'cell' || !values.length) return;
    const table = this.doc.shapeById(selection.slideIndex, selection.shapeId);
    if (!table) return;
    const block = tableSelectionBlock(selection);
    if (!canPasteTableCells(table, block.row, block.col, values)) {
      this.toast('error', t('Split merged cells before pasting multiple cells'));
      return;
    }
    this.doc.transact(t('Paste table cells'), () => {
      pasteTableCells(table, block.row, block.col, values);
      this.doc.selectCell(selection.slideIndex, selection.shapeId, block.row, block.col);
      this.doc.selectCell(
        selection.slideIndex,
        selection.shapeId,
        block.row + values.length - 1,
        block.col + values.reduce((width, row) => Math.max(width, row.length), 0) - 1,
        true,
      );
    });
  }

  hasClipboard(): boolean {
    return this.#clipboard != null;
  }

  exitGroup(): boolean {
    const slide = this.doc.currentSlide;
    const parent = slide && shapeScope(slide, selectedShapeId(this.doc.selection)).parent;
    if (!parent) return false;
    this.doc.selectShape(this.doc.selection.slideIndex, getShapeId(parent));
    return true;
  }

  /** Move all selected shapes by a slide-space EMU delta as one undo step. */
  nudge(dxEmu: number, dyEmu: number): void {
    if (this.selectionLocked()) return;
    const shapes = this.selectedShapes();
    if (!shapes.length) return;
    const slide = this.doc.currentSlide;
    if (!slide) return;
    const inverse = invert(shapeScope(slide, selectedShapeId(this.doc.selection)).matrix);
    if (!inverse) return;
    const origin = project(inverse, { x: 0, y: 0 });
    const end = project(inverse, { x: dxEmu, y: dyEmu });
    dxEmu = end.x - origin.x;
    dyEmu = end.y - origin.y;
    this.doc.transact(t('Move'), () => {
      for (const s of shapes) {
        const b = getShapeBoundsResolved(this.doc.pres, s);
        if (!b) continue;
        setShapeBounds(s, {
          x: emu(b.x + dxEmu),
          y: emu(b.y + dyEmu),
          w: b.w,
          h: b.h,
        });
      }
    });
  }
}
