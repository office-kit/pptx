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
import { availableOperands, selectedShapeId } from './selection.ts';

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
      const slides = doc.slides;
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

const registry = new Map<string, Command>(
  capabilities.map((cap) => [cap.id, new ManifestCommand(cap)]),
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
