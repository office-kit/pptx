// The single source of truth for the editor: the in-memory presentation, the
// current selection, undo/redo history, and derived rendering.
//
// The @office-kit/pptx model is a mutable object graph that the library's
// authoring functions edit in place. Rather than deep-proxy that graph (which
// fights the library's reference-based mutation), we hold it in `$state.raw`
// and drive reactivity with an explicit `version` counter bumped on every
// mutation. Derived values (the rendered SVG, the slide list) read `version`,
// so any command re-renders the canvas without the store needing to understand
// the shape of every mutation.
//
// Undo/redo snapshots the presentation by *serializing it to .pptx bytes*, not
// by structuredClone: the model's real state hangs off a symbol-keyed
// `OpcPackage` instance that structuredClone silently drops, producing a
// corrupt copy. save→load is the library's guaranteed round-trip, so it is the
// only sound snapshot. Edits stay instant (the mutation runs synchronously);
// savePresentation captures bytes synchronously and returns a promise. Store
// that promise immediately so rapid edits remain distinct. Because each discrete
// gesture (a drag, a click) commits exactly one transaction, this is one
// serialization per user action — not per pointer move.

import {
  createPresentation,
  setGridSpacing,
  setSnapToGrid,
  addTitleSlide,
  getSlides,
  findShapeById,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import type {
  PresentationData,
  SlideData,
  SlideShapeData,
  Color,
  GradientFillOptions,
} from '@office-kit/pptx';
import { RegroupHistory } from './regroup-history.ts';
import { selectedSlideIndices, type Selection } from './selection.ts';

interface Snapshot {
  readonly bytes: Promise<Uint8Array>;
  readonly selection: Selection;
  readonly label: string;
  readonly formerGroups: RegroupHistory['records'];
}

const HISTORY_MAX = 60;

export class EditorDocument {
  readonly rememberedFills = new Map<
    string,
    { solid?: { color: Color; opacity?: number }; gradient?: GradientFillOptions }
  >();
  /** The live presentation. Mutated in place by library commands. */
  pres = $state.raw<PresentationData>(createInitial());
  /** Bumped on every mutation to invalidate derived rendering. */
  version = $state(0);
  committedVersion = $state(0);
  liveEditing = $state(false);
  selection = $state.raw<Selection>({ kind: 'none', slideIndex: 0 });
  fileName = $state<string>('Untitled.pptx');
  dirty = $state<boolean>(false);

  // History of committed states as serialized snapshots. `#cursor` is the index
  // of the current state within `#history`.
  #history = $state.raw<Snapshot[]>([]);
  #cursor = $state(-1);
  // A requested restore may still be loading while another undo is requested.
  #requestedCursor = $state(-1);
  #operation = 0;
  readonly regroupHistory = new RegroupHistory();

  constructor() {
    // Seed the initial state so the first undo returns to the blank deck.
    this.#snapshot('Initial');
  }

  // --- Derived views -----------------------------------------------------
  slides = $derived.by<ReadonlyArray<SlideData>>(() => {
    this.version;
    return getSlides(this.pres);
  });

  currentSlide = $derived.by<SlideData | null>(() => {
    const list = this.slides;
    const idx = this.selection.slideIndex;
    return list[idx] ?? list[0] ?? null;
  });

  currentSvg = $derived.by<string>(() => {
    this.version;
    const slide = this.currentSlide;
    if (!slide) return '';
    try {
      return renderSlideToSvg(this.pres, slide);
    } catch (err) {
      return `<!-- render error: ${(err as Error).message} -->`;
    }
  });

  canUndo = $derived(this.#requestedCursor > 0);
  canRedo = $derived(
    this.#requestedCursor >= 0 && this.#requestedCursor < this.#history.length - 1,
  );

  // --- Resolvers ---------------------------------------------------------
  slideAt(index: number): SlideData | null {
    return this.slides[index] ?? null;
  }

  shapeById(slideIndex: number, id: number): SlideShapeData | null {
    const slide = this.slideAt(slideIndex);
    return slide ? findShapeById(slide, id) : null;
  }

  // --- Mutation core -----------------------------------------------------
  /**
   * Run `fn` as one undoable transaction. The mutation runs synchronously so
   * the UI updates immediately; a byte snapshot is captured asynchronously for
   * undo. Returns whatever `fn` returns (e.g. a newly created shape/slide).
   */
  transact<T>(label: string, fn: () => T): T {
    this.#invalidateRestore();
    const current = this.#history[this.#cursor];
    if (current)
      this.#history = this.#history.map((snapshot, index) =>
        index === this.#cursor ? { ...snapshot, selection: this.selection } : snapshot,
      );
    const result = fn();
    this.version++;
    this.dirty = true;
    this.#snapshot(label);
    return result;
  }

  /**
   * Apply a live, high-frequency mutation (a drag/resize frame). Re-renders
   * immediately but does NOT snapshot — so a whole gesture stays one undo step.
   * Call `commit()` once when the gesture ends.
   */
  applyLive<T>(fn: () => T): T {
    this.liveEditing = true;
    this.#invalidateRestore();
    const result = fn();
    this.version++;
    this.dirty = true;
    return result;
  }

  /** Close a live gesture by taking a single undo snapshot. */
  commit(label: string): void {
    this.liveEditing = false;
    this.#snapshot(label);
  }

  #invalidateRestore(): void {
    this.#operation++;
    this.#requestedCursor = this.#cursor;
  }

  /** Capture before yielding: neither the document nor selection may drift. */
  #snapshot(label: string): void {
    this.committedVersion = this.version;
    const bytes = savePresentation(this.pres);
    const kept = this.#history.slice(0, this.#cursor + 1);
    this.regroupHistory.prune(getSlides(this.pres));
    kept.push({
      bytes,
      selection: this.selection,
      label,
      formerGroups: this.regroupHistory.records,
    });
    this.#history = kept.slice(-HISTORY_MAX);
    this.#cursor = this.#history.length - 1;
    this.#requestedCursor = this.#cursor;
  }

  async #restore(index: number, selection?: Selection): Promise<void> {
    const snap = this.#history[index];
    if (!snap) return;
    const operation = ++this.#operation;
    this.#requestedCursor = index;
    try {
      const pres = await loadPresentation(await snap.bytes);
      // New/Open, another restore, or an edit takes precedence over stale work.
      if (operation !== this.#operation) return;
      this.pres = pres;
      this.regroupHistory.records = snap.formerGroups;
      this.selection = selection ?? snap.selection;
      this.liveEditing = false;
      this.#cursor = index;
      this.version++;
      this.committedVersion = this.version;
      this.dirty = true;
    } catch (error) {
      if (operation === this.#operation) this.#requestedCursor = this.#cursor;
      throw error;
    }
  }

  /** Discard the current gesture without adding an undo step or dropping redo. */
  async cancelLive(): Promise<void> {
    if (this.liveEditing) await this.#restore(this.#cursor, this.selection);
  }

  async undo(): Promise<void> {
    if (this.#requestedCursor > 0) await this.#restore(this.#requestedCursor - 1);
  }

  async redo(): Promise<void> {
    if (this.#requestedCursor < this.#history.length - 1)
      await this.#restore(this.#requestedCursor + 1);
  }

  // --- Selection ---------------------------------------------------------
  select(sel: Selection): void {
    this.selection = sel;
  }

  selectSlide(index: number, options: { additive?: boolean; range?: boolean } = {}): void {
    const clamped = Math.max(0, Math.min(index, getSlides(this.pres).length - 1));
    const current = this.selection;
    const anchor = current.kind === 'slide' ? (current.anchorIndex ?? current.slideIndex) : clamped;
    let indices = [clamped];
    if (options.range) {
      indices = Array.from(
        { length: Math.abs(clamped - anchor) + 1 },
        (_, i) => Math.min(clamped, anchor) + i,
      );
    } else if (options.additive && current.kind === 'slide') {
      const selected = new Set(selectedSlideIndices(current));
      if (selected.has(clamped) && selected.size > 1) selected.delete(clamped);
      else selected.add(clamped);
      indices = [...selected].sort((a, b) => a - b);
    }
    this.selection = {
      kind: 'slide',
      slideIndex: indices.includes(clamped) ? clamped : indices[indices.length - 1]!,
      slideIndices: indices,
      anchorIndex: options.range ? anchor : clamped,
    };
  }

  selectShape(slideIndex: number, shapeId: number, additive = false): void {
    if (additive && this.selection.kind === 'shape' && this.selection.slideIndex === slideIndex) {
      const set = new Set(this.selection.shapeIds);
      if (set.has(shapeId)) set.delete(shapeId);
      else set.add(shapeId);
      this.selection = { kind: 'shape', slideIndex, shapeIds: [...set] };
    } else {
      this.selection = { kind: 'shape', slideIndex, shapeIds: [shapeId] };
    }
  }

  selectCell(slideIndex: number, shapeId: number, row: number, col: number, extend = false): void {
    const current = this.selection;
    if (
      extend &&
      current.kind === 'cell' &&
      current.slideIndex === slideIndex &&
      current.shapeId === shapeId
    ) {
      this.selection = { ...current, end: { row, col } };
      return;
    }
    this.selection = { kind: 'cell', slideIndex, shapeId, row, col };
  }

  clearShapeSelection(): void {
    this.selection = { kind: 'none', slideIndex: this.selection.slideIndex };
  }

  // --- IO ----------------------------------------------------------------
  async loadBytes(bytes: Uint8Array, name: string): Promise<void> {
    this.#invalidateRestore();
    const operation = this.#operation;
    const pres = await loadPresentation(bytes);
    if (operation !== this.#operation) return;
    this.pres = pres;
    this.fileName = name;
    this.regroupHistory.records = [];
    this.rememberedFills.clear();
    this.#history = [];
    this.#cursor = -1;
    this.selection = { kind: 'none', slideIndex: 0 };
    this.version++;
    this.dirty = false;
    this.liveEditing = false;
    this.#snapshot('Open');
  }

  async toBytes(): Promise<Uint8Array> {
    return savePresentation(this.pres);
  }

  markSaved(version: number): void {
    if (version === this.version) this.dirty = false;
  }

  resetBlank(): void {
    this.#invalidateRestore();
    this.pres = createInitial();
    this.fileName = 'Untitled.pptx';
    this.regroupHistory.records = [];
    this.#history = [];
    this.#cursor = -1;
    this.selection = { kind: 'none', slideIndex: 0 };
    this.version++;
    this.dirty = false;
    this.liveEditing = false;
    this.#snapshot('New');
  }
}

function createInitial(): PresentationData {
  const pres = createPresentation();
  addTitleSlide(pres, 'Untitled presentation');
  try {
    const defaults = JSON.parse(localStorage.getItem('office-grid-defaults') ?? 'null');
    if (
      defaults &&
      Number.isSafeInteger(defaults.x) &&
      defaults.x > 0 &&
      Number.isSafeInteger(defaults.y) &&
      defaults.y > 0 &&
      typeof defaults.snap === 'boolean'
    ) {
      setGridSpacing(pres, { x: defaults.x, y: defaults.y });
      setSnapToGrid(pres, defaults.snap);
    }
  } catch {
    /* Missing or invalid local defaults do not prevent creating a deck. */
  }
  return pres;
}
