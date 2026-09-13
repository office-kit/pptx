// The editor controller — the façade the UI talks to.
//
// It owns the document, runs commands through the registry, and holds the
// transient UI state that several components share (which command dialog is
// open, whether the palette is showing, the toast queue). Ribbon buttons,
// context menus and the palette all funnel through `invoke` / `runOrPrompt`,
// so there is exactly one path from "user intent" to "library call".

import {
  copyShape,
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

/** What copy/cut stashed: the slide + shape ids to clone on paste. */
interface Clipboard {
  slideIndex: number;
  shapeIds: number[];
}

const PASTE_OFFSET = inches(0.25) as unknown as number;

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
  #clipboard: Clipboard | null = null;

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
    const shapes = this.selectedShapes();
    if (!shapes.length) return;
    this.doc.transact('Delete', () => {
      for (const s of shapes) removeShape(s);
    });
    this.doc.clearShapeSelection();
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
          x: ((b.x as unknown as number) + offset) as never,
          y: ((b.y as unknown as number) + offset) as never,
          w: b.w,
          h: b.h,
        });
      }
      newIds.push(getShapeId(copy));
    }
    return newIds;
  }

  duplicateSelection(): void {
    const shapes = this.selectedShapes();
    if (!shapes.length) return;
    const slideIndex = this.doc.selection.slideIndex;
    const newIds = this.doc.transact('Duplicate', () =>
      this.#cloneOnto(shapes, slideIndex, PASTE_OFFSET),
    );
    if (newIds.length) this.doc.select({ kind: 'shape', slideIndex, shapeIds: newIds });
  }

  copySelection(): void {
    const sel = this.doc.selection;
    const ids = selectedShapeIds(sel);
    if (!ids.length) return;
    this.#clipboard = { slideIndex: sel.slideIndex, shapeIds: [...ids] };
    this.toast('info', `Copied ${ids.length} shape${ids.length > 1 ? 's' : ''}`);
  }

  cutSelection(): void {
    this.copySelection();
    this.deleteSelection();
  }

  paste(): void {
    const clip = this.#clipboard;
    if (!clip) return;
    const sources = clip.shapeIds
      .map((id) => this.doc.shapeById(clip.slideIndex, id))
      .filter((s): s is SlideShapeData => s != null);
    if (!sources.length) return;
    const slideIndex = this.doc.selection.slideIndex;
    const newIds = this.doc.transact('Paste', () =>
      this.#cloneOnto(sources, slideIndex, PASTE_OFFSET),
    );
    if (newIds.length) this.doc.select({ kind: 'shape', slideIndex, shapeIds: newIds });
  }

  hasClipboard(): boolean {
    return this.#clipboard != null;
  }

  /** Move all selected shapes by an EMU delta as one undo step. */
  nudge(dxEmu: number, dyEmu: number): void {
    const shapes = this.selectedShapes();
    if (!shapes.length) return;
    this.doc.transact('Move', () => {
      for (const s of shapes) {
        const b = getShapeBoundsResolved(this.doc.pres, s);
        if (!b) continue;
        setShapeBounds(s, {
          x: ((b.x as unknown as number) + dxEmu) as never,
          y: ((b.y as unknown as number) + dyEmu) as never,
          w: b.w,
          h: b.h,
        });
      }
    });
  }
}
