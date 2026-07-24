// The editor selection model.
//
// A capability's `operand` (presentation / slide / shape / cell) determines
// which selection makes it applicable — this mirrors PowerPoint's contextual
// behaviour, where the ribbon lights up different controls depending on whether
// you have a slide, a shape, or a table cell selected.
//
// Selections are stored by *stable identity* (slide index + shape id + cell
// coordinates), never by object reference, so they survive re-renders, undo
// snapshots, and structural clones.

import type { Operand } from '../manifest/types.ts';

export interface SlideSelection {
  readonly kind: 'slide';
  readonly slideIndex: number;
}

export interface ShapeSelection {
  readonly kind: 'shape';
  readonly slideIndex: number;
  /** `getShapeId` value; unique within the slide. Multiple for multi-select. */
  readonly shapeIds: readonly number[];
}

export interface CellSelection {
  readonly kind: 'cell';
  readonly slideIndex: number;
  /** The table shape holding the cell. */
  readonly shapeId: number;
  readonly row: number;
  readonly col: number;
}

export interface NoneSelection {
  readonly kind: 'none';
  /** Even with nothing selected we track which slide is shown. */
  readonly slideIndex: number;
}

export type Selection = SlideSelection | ShapeSelection | CellSelection | NoneSelection;

export function selectionSlideIndex(sel: Selection): number {
  return sel.slideIndex;
}

/** The primary (first) selected shape id, if any. */
export function selectedShapeId(sel: Selection): number | null {
  if (sel.kind === 'shape') return sel.shapeIds[0] ?? null;
  if (sel.kind === 'cell') return sel.shapeId;
  return null;
}

/** All selected shape ids (empty unless a shape/cell selection). */
export function selectedShapeIds(sel: Selection): number[] {
  if (sel.kind === 'shape') return [...sel.shapeIds];
  if (sel.kind === 'cell') return [sel.shapeId];
  return [];
}

/** Which operands the current selection can satisfy. Used to enable/disable
 *  commands and to route a command to the right target object. */
export function availableOperands(sel: Selection): ReadonlySet<Operand> {
  const set = new Set<Operand>(['presentation']);
  // A presentation always has slides once one exists.
  set.add('slide');
  if (sel.kind === 'shape' && sel.shapeIds.length > 0) set.add('shape');
  if (sel.kind === 'cell') {
    set.add('shape');
    set.add('cell');
  }
  return set;
}

export function isOperandAvailable(sel: Selection, operand: Operand): boolean {
  return availableOperands(sel).has(operand);
}
