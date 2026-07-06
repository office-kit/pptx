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
// the byte snapshot is taken asynchronously afterwards, which is why undo
// availability can trail an edit by a few milliseconds. Because each discrete
// gesture (a drag, a click) commits exactly one transaction, this is one
// serialization per user action — not per pointer move.

import {
  createPresentation,
  addTitleSlide,
  getSlides,
  findShapeById,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import type { PresentationData, SlideData, SlideShapeData } from '@office-kit/pptx';
import type { Selection } from './selection.ts';

interface Snapshot {
  readonly bytes: Uint8Array;
  readonly selection: Selection;
  readonly label: string;
}

const HISTORY_MAX = 60;

export class EditorDocument {
  /** The live presentation. Mutated in place by library commands. */
  pres = $state.raw<PresentationData>(createInitial());
  /** Bumped on every mutation to invalidate derived rendering. */
  version = $state(0);
  selection = $state.raw<Selection>({ kind: 'none', slideIndex: 0 });
  fileName = $state<string>('Untitled.pptx');
  dirty = $state<boolean>(false);

  // History of committed states as serialized snapshots. `#cursor` is the index
  // of the current state within `#history`.
  #history = $state.raw<Snapshot[]>([]);
  #cursor = $state(-1);
  #serializing = false;
  #pending = false;

  constructor() {
    // Seed the initial state so the first undo returns to the blank deck.
    void this.#snapshot('Initial');
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

  canUndo = $derived(this.#cursor > 0);
  canRedo = $derived(this.#cursor >= 0 && this.#cursor < this.#history.length - 1);

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
    const result = fn();
    this.version++;
    this.dirty = true;
    void this.#snapshot(label);
    return result;
  }

  /**
   * Apply a live, high-frequency mutation (a drag/resize frame). Re-renders
   * immediately but does NOT snapshot — so a whole gesture stays one undo step.
   * Call `commit()` once when the gesture ends.
   */
  applyLive<T>(fn: () => T): T {
    const result = fn();
    this.version++;
    this.dirty = true;
    return result;
  }

  /** Close a live gesture by taking a single undo snapshot. */
  commit(label: string): void {
    void this.#snapshot(label);
  }

  /** Serialize the current state and push it as the new history head. */
  async #snapshot(label: string): Promise<void> {
    if (this.#serializing) {
      this.#pending = true;
      return;
    }
    this.#serializing = true;
    try {
      const bytes = await savePresentation(this.pres);
      const kept = this.#history.slice(0, this.#cursor + 1);
      kept.push({ bytes, selection: this.selection, label });
      const trimmed = kept.slice(-HISTORY_MAX);
      this.#history = trimmed;
      this.#cursor = trimmed.length - 1;
    } finally {
      this.#serializing = false;
      if (this.#pending) {
        this.#pending = false;
        void this.#snapshot(label);
      }
    }
  }

  async #restore(index: number): Promise<void> {
    const snap = this.#history[index];
    if (!snap) return;
    this.pres = await loadPresentation(snap.bytes);
    this.selection = snap.selection;
    this.#cursor = index;
    this.version++;
    this.dirty = true;
  }

  async undo(): Promise<void> {
    if (this.#cursor > 0) await this.#restore(this.#cursor - 1);
  }

  async redo(): Promise<void> {
    if (this.#cursor < this.#history.length - 1) await this.#restore(this.#cursor + 1);
  }

  // --- Selection ---------------------------------------------------------
  select(sel: Selection): void {
    this.selection = sel;
  }

  selectSlide(index: number): void {
    const clamped = Math.max(0, Math.min(index, this.slides.length - 1));
    this.selection = { kind: 'none', slideIndex: clamped };
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

  selectCell(slideIndex: number, shapeId: number, row: number, col: number): void {
    this.selection = { kind: 'cell', slideIndex, shapeId, row, col };
  }

  clearShapeSelection(): void {
    this.selection = { kind: 'none', slideIndex: this.selection.slideIndex };
  }

  // --- IO ----------------------------------------------------------------
  async loadBytes(bytes: Uint8Array, name: string): Promise<void> {
    const pres = await loadPresentation(bytes);
    this.pres = pres;
    this.fileName = name;
    this.#history = [];
    this.#cursor = -1;
    this.selection = { kind: 'none', slideIndex: 0 };
    this.version++;
    this.dirty = false;
    void this.#snapshot('Open');
  }

  async toBytes(): Promise<Uint8Array> {
    return savePresentation(this.pres);
  }

  resetBlank(): void {
    this.pres = createInitial();
    this.fileName = 'Untitled.pptx';
    this.#history = [];
    this.#cursor = -1;
    this.selection = { kind: 'none', slideIndex: 0 };
    this.version++;
    this.dirty = false;
    void this.#snapshot('New');
  }
}

function createInitial(): PresentationData {
  const pres = createPresentation();
  addTitleSlide(pres, 'Untitled presentation');
  return pres;
}
