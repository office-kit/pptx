// The command registry — where the capability manifest becomes executable.
//
// Every capability in the manifest is turned into a Command that:
//   - knows its operand and can resolve it from the live selection,
//   - reports whether it can run given the current selection, and
//   - executes by calling the *real* library function by name, inside an
//     undoable transaction, and post-processes the result (e.g. selecting a
//     newly created shape).
//
// Because the registry is built by iterating the manifest, coverage is
// structural: if a capability is manifested (and the coverage test guarantees
// all 147 are), it has a runnable command here — reachable at minimum through
// the command palette. Bespoke ribbon UIs supply nicer argument collection but
// dispatch through this same execution path.

import * as pptx from '@office-kit/pptx';
import type { PresentationData, SlideData, SlideShapeData } from '@office-kit/pptx';
import { capabilities, capabilityById } from '../manifest/index.ts';
import type { ResolvedCapability } from '../manifest/types.ts';
import type { Selection } from './selection.ts';
import { availableOperands, selectedShapeId, topLevelShapes } from './selection.ts';

/** A dynamic view of the library so we can dispatch by capability id. */
const lib = pptx as unknown as Record<string, (...args: unknown[]) => unknown>;

/**
 * The document surface the registry needs — the structural subset of
 * `EditorDocument` that command dispatch touches. Declared here (not imported
 * from `document.svelte.ts`) so the registry — and the coverage/smoke tests that
 * import it — stay free of Svelte-rune modules, which the root `tsc` can't parse.
 */
export interface CommandDoc {
  readonly selection: Selection;
  readonly pres: PresentationData;
  readonly slides: ReadonlyArray<SlideData>;
  slideAt(index: number): SlideData | null;
  shapeById(slideIndex: number, id: number): SlideShapeData | null;
  select(selection: Selection): void;
  selectShape(slideIndex: number, id: number): void;
  selectSlide(index: number): void;
  transact<T>(label: string, fn: () => T): T;
}

export interface CommandContext {
  readonly doc: CommandDoc;
}

/** The concrete target object a capability operates on, resolved from selection. */
export interface CellTarget {
  readonly table: unknown;
  readonly row: number;
  readonly col: number;
}

export class CommandError extends Error {}

export interface Command {
  readonly capability: ResolvedCapability;
  /** Args this command still needs from the user (operand param dropped). */
  readonly params: ResolvedCapability['params'];
  /** True if the current selection can supply this command's operand. */
  canRun(ctx: CommandContext): boolean;
  /**
   * Execute with the given named arguments. Runs inside `doc.transact`, calls
   * the library function, and applies result-selection side effects.
   */
  run(ctx: CommandContext, args: Record<string, unknown>): unknown;
}

/** Resolve the object a capability's first parameter expects, from selection. */
function resolveOperand(doc: CommandDoc, cap: ResolvedCapability): unknown {
  const sel = doc.selection;
  switch (cap.operand) {
    case 'presentation':
      return doc.pres;
    case 'slide':
      return doc.slideAt(sel.slideIndex);
    case 'shape': {
      const id = selectedShapeId(sel);
      return id == null ? null : doc.shapeById(sel.slideIndex, id);
    }
    case 'cell': {
      if (sel.kind !== 'cell') return null;
      const table = doc.shapeById(sel.slideIndex, sel.shapeId);
      if (!table) return null;
      try {
        return (lib.getTableCell as (t: unknown, r: number, c: number) => unknown)(
          table,
          sel.row,
          sel.col,
        );
      } catch {
        return null;
      }
    }
  }
}

/** Order the user-supplied named args to positional args. `cap.params` already
 *  excludes the operand, so this maps every entry. */
function orderArgs(cap: ResolvedCapability, args: Record<string, unknown>): unknown[] {
  return cap.params.map((p) => args[p.name]);
}

/** After a mutating call, if it produced a shape/slide, select it. */
function applyResultSelection(doc: CommandDoc, cap: ResolvedCapability, result: unknown): void {
  if (result == null || typeof result !== 'object') return;
  const ret = cap.returns;
  try {
    if (ret.includes('SlideShapeData')) {
      const id = (lib.getShapeId as (s: unknown) => number)(result);
      doc.selectShape(doc.selection.slideIndex, id);
    } else if (ret.includes('SlideData')) {
      const slides = pptx.getSlides(doc.pres);
      const idx = slides.indexOf(result as never);
      if (idx >= 0) doc.selectSlide(idx);
    }
  } catch {
    // Result-selection is a convenience; never fail the command over it.
  }
}

class ManifestCommand implements Command {
  readonly capability: ResolvedCapability;
  constructor(cap: ResolvedCapability) {
    this.capability = cap;
  }

  get params(): ResolvedCapability['params'] {
    return this.capability.params;
  }

  canRun(ctx: CommandContext): boolean {
    const cap = this.capability;
    if (!cap.takesOperand) return true; // factory/package ops always available
    if (!availableOperands(ctx.doc.selection).has(cap.operand)) return false;
    return resolveOperand(ctx.doc, cap) != null || cap.operand === 'presentation';
  }

  run(ctx: CommandContext, args: Record<string, unknown>): unknown {
    const cap = this.capability;
    const fn = lib[cap.id];
    if (typeof fn !== 'function') {
      throw new CommandError(`Library function "${cap.id}" is not callable.`);
    }
    const positional = orderArgs(cap, args);
    return ctx.doc.transact(cap.labelEn, () => {
      let result: unknown;
      if (cap.takesOperand) {
        const operand = resolveOperand(ctx.doc, cap);
        if (operand == null && cap.operand !== 'presentation') {
          throw new CommandError(`No ${cap.operand} selected for "${cap.id}".`);
        }
        result = fn(operand, ...positional);
      } else {
        result = fn(...positional);
      }
      applyResultSelection(ctx.doc, cap, result);
      return result;
    });
  }
}

// Slide operations need both the presentation and the active slide. Keep these
// bindings here so ribbon, palette and navigator expose the same commands.
const activeSlideCommands = new Set([
  'duplicateSlide',
  'removeSlide',
  'moveSlide',
  'duplicateSlideAt',
]);
class SlideCommand extends ManifestCommand {
  override get params(): ResolvedCapability['params'] {
    return super.params.filter((param) => param.name !== 'slide');
  }

  override canRun(ctx: CommandContext): boolean {
    return (
      this.capability.id === 'addBlankSlide' ||
      ctx.doc.slideAt(ctx.doc.selection.slideIndex) !== null
    );
  }

  override run(ctx: CommandContext, args: Record<string, unknown>): unknown {
    const doc = ctx.doc;
    const index = doc.selection.slideIndex;
    const slide = doc.slideAt(index);
    const id = this.capability.id;
    if (id !== 'addBlankSlide' && !slide) throw new CommandError('No slide selected.');
    return doc.transact(this.capability.labelEn, () => {
      switch (id) {
        case 'addBlankSlide': {
          const added = pptx.addBlankSlide(doc.pres);
          pptx.moveSlide(doc.pres, added, index + 1);
          doc.selectSlide(Math.min(index + 1, pptx.getSlides(doc.pres).length - 1));
          return pptx.getSlides(doc.pres)[doc.selection.slideIndex];
        }
        case 'duplicateSlide':
        case 'duplicateSlideAt': {
          const at = id === 'duplicateSlide' ? index + 1 : args.atIndex;
          if (typeof at !== 'number' || !Number.isInteger(at))
            throw new CommandError('Slide position must be an integer.');
          const duplicate = pptx.duplicateSlideAt(doc.pres, at, slide!);
          doc.selectSlide(pptx.getSlides(doc.pres).indexOf(duplicate));
          return duplicate;
        }
        case 'removeSlide':
          pptx.removeSlide(doc.pres, slide!);
          doc.selectSlide(index);
          return;
        case 'moveSlide': {
          const at = args.toIndex;
          if (typeof at !== 'number' || !Number.isInteger(at))
            throw new CommandError('Slide position must be an integer.');
          pptx.moveSlide(doc.pres, slide!, at);
          doc.selectSlide(at);
          return;
        }
        default:
          throw new CommandError(`Unknown slide command: ${id}`);
      }
    });
  }
}

/** Group commands consume the selection, never a JSON representation of shapes. */
class GroupCommand extends ManifestCommand {
  override get params(): ResolvedCapability['params'] {
    return super.params.filter((param) => param.name !== 'shapes');
  }

  private shapes(doc: CommandDoc): SlideShapeData[] {
    const selection = doc.selection;
    const slide = doc.slideAt(selection.slideIndex);
    if (!slide || selection.kind !== 'shape') return [];
    const ids = new Set(selection.shapeIds);
    // A user can select front to back. Keep the existing stacking order.
    return topLevelShapes(slide).filter((shape) => ids.has(pptx.getShapeId(shape)));
  }

  override canRun({ doc }: CommandContext): boolean {
    const shapes = this.shapes(doc);
    return this.capability.id === 'groupShapes'
      ? shapes.length >= 2
      : shapes.some((shape) => pptx.getShapeKind(shape) === 'group');
  }

  override run({ doc }: CommandContext, args: Record<string, unknown>): unknown {
    if (!this.canRun({ doc }))
      throw new CommandError('Select shapes to group or a group to ungroup.');
    const shapes = this.shapes(doc);
    let name: string | undefined;
    if (args.opts !== undefined) {
      if (typeof args.opts !== 'object' || args.opts === null)
        throw new CommandError('Group options must be an object.');
      if ('name' in args.opts) {
        if (typeof args.opts.name !== 'string') throw new CommandError('Group name must be text.');
        name = args.opts.name;
      }
    }
    return doc.transact(this.capability.labelEn, () => {
      if (this.capability.id === 'groupShapes') {
        const group = pptx.groupShapes(shapes, name === undefined ? {} : { name });
        doc.selectShape(doc.selection.slideIndex, pptx.getShapeId(group));
        return group;
      }
      const children = shapes.flatMap((shape) =>
        pptx.getShapeKind(shape) === 'group' ? [...pptx.ungroupShapes(shape)] : [shape],
      );
      doc.select({
        kind: 'shape',
        slideIndex: doc.selection.slideIndex,
        shapeIds: children.map(pptx.getShapeId),
      });
      return children;
    });
  }
}

const registry = new Map<string, Command>(
  capabilities.map((cap) => [
    cap.id,
    activeSlideCommands.has(cap.id) || cap.id === 'addBlankSlide'
      ? new SlideCommand(cap)
      : cap.id === 'groupShapes' || cap.id === 'ungroupShapes'
        ? new GroupCommand(cap)
        : new ManifestCommand(cap),
  ]),
);

export function getCommand(id: string): Command | undefined {
  return registry.get(id);
}

export function allCommands(): readonly Command[] {
  return [...registry.values()];
}

/** Commands the current selection can run right now. */
export function runnableCommands(ctx: CommandContext): readonly Command[] {
  return allCommands().filter((c) => c.canRun(ctx));
}

export function hasCommand(id: string): boolean {
  return registry.has(id);
}

// Re-export so UI can introspect the manifest via the registry entry point.
export { capabilityById };
