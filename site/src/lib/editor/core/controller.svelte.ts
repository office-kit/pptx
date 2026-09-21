// The editor controller — the façade the UI talks to.
//
// It owns the document, runs commands through the registry, and holds the
// transient UI state that several components share (which command dialog is
// open, whether the palette is showing, the toast queue). Ribbon buttons,
// context menus and the palette all funnel through `invoke` / `runOrPrompt`,
// so there is exactly one path from "user intent" to "library call".

import {
  copyShape,
  emu,
  loadPresentation,
  getSlides,
  findShapeById,
  type PresentationData,
  getShapeBoundsResolved,
  getShapeId,
  getSlideShapes,
  inches,
  removeShape,
  setShapeBounds,
  type SlideShapeData,
} from '@office-kit/pptx';
import { getCommand, type Command, type CommandContext } from './registry.ts';
import { capabilityById } from '../manifest/index.ts';
import { EditorDocument } from './document.svelte.ts';
import { t } from '../i18n/i18n.svelte.ts';
import { selectedShapeIds } from './selection.ts';

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
  shapeIds: number[];
}

const PASTE_OFFSET = inches(0.25);

let toastSeq = 0;

export class EditorController {
  readonly doc = new EditorDocument();

  /** Command whose argument dialog is currently open (null = none). */
  activeDialog = $state<string | null>(null);
  paletteOpen = $state<boolean>(false);
  toasts = $state<Toast[]>([]);

  get ctx(): CommandContext {
    return { doc: this.doc };
  }

  command(id: string): Command | undefined {
    return getCommand(id);
  }

  canRun(id: string): boolean {
    const cmd = getCommand(id);
    return cmd ? cmd.canRun(this.ctx) : false;
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

  /** Args seeded into a dialog opened via runOrPrompt. */
  pendingPreset: Record<string, unknown> = {};

  closeDialog(): void {
    this.activeDialog = null;
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
  /** When set by the canvas, `fit` recomputes to this multiplier. */
  fitZoom = $state(1);

  setZoom(z: number): void {
    this.zoom = Math.max(0.1, Math.min(z, 5));
  }
  zoomIn(): void {
    this.setZoom(this.zoom * 1.2);
  }
  zoomOut(): void {
    this.setZoom(this.zoom / 1.2);
  }
  zoomFit(): void {
    this.setZoom(this.fitZoom);
  }
  zoomReset(): void {
    this.setZoom(1);
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
  #clipboard = $state.raw<Clipboard | null>(null);

  /** Resolve the currently selected shapes to live objects. */
  selectedShapes(): SlideShapeData[] {
    const sel = this.doc.selection;
    const ids = selectedShapeIds(sel);
    return ids
      .map((id) => this.doc.shapeById(sel.slideIndex, id))
      .filter((s): s is SlideShapeData => s != null);
  }

  selectAllShapes(): void {
    const slideIndex = this.doc.selection.slideIndex;
    const slide = this.doc.slideAt(slideIndex);
    if (!slide) return;
    const ids = getSlideShapes(slide).map((s) => getShapeId(s));
    if (ids.length) this.doc.select({ kind: 'shape', slideIndex, shapeIds: ids });
  }

  deleteSelection(): void {
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
      const copy = copyShape(slide, src);
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

  copySelection(): void {
    const sel = this.doc.selection;
    const ids = selectedShapeIds(sel);
    if (!ids.length) return;
    this.#clipboard = {
      presentation: this.doc.toBytes().then(loadPresentation),
      slideIndex: sel.slideIndex,
      shapeIds: [...ids],
    };
    // Attach a handler immediately; a later paste reports the captured failure.
    void this.#clipboard.presentation.catch((error: Error) => this.toast('error', error.message));
  }

  cutSelection(): void {
    if (!this.selectedShapes().length) return;
    this.copySelection();
    this.deleteSelection();
  }

  async paste(): Promise<void> {
    const clip = this.#clipboard;
    if (!clip) return;
    const targetPresentation = this.doc.pres;
    const targetSlide = this.doc.slideAt(this.doc.selection.slideIndex);
    if (!targetSlide) return;
    try {
      const source = await clip.presentation;
      // New/Open or undo replaces the document. Never paste into a stale target.
      if (this.doc.pres !== targetPresentation) return;
      const slideIndex = getSlides(targetPresentation).indexOf(targetSlide);
      if (slideIndex < 0) return;
      const sourceSlide = getSlides(source)[clip.slideIndex]!;
      const sources = clip.shapeIds.map((id) => findShapeById(sourceSlide, id)!);
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

  hasClipboard(): boolean {
    return this.#clipboard != null;
  }

  /** Move all selected shapes by an EMU delta as one undo step. */
  nudge(dxEmu: number, dyEmu: number): void {
    const shapes = this.selectedShapes();
    if (!shapes.length) return;
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
