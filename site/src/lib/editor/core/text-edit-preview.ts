import {
  addBlankSlide,
  copyShape,
  createPresentation,
  getShapeText,
  getTableCells,
  getTableCellText,
  setShapeText,
  setTableCellText,
  type SlideShapeData,
} from '@office-kit/pptx';

type TextEdit = { start: number; end: number; text: string };
type CellPosition = { row: number; col: number };

/** Use the same formatting-preserving edits for preview and commit. */
export function replayTextEdits(
  shape: SlideShapeData,
  changes: readonly TextEdit[],
  position?: CellPosition,
): void {
  const cell = position ? getTableCells(shape)[position.row]![position.col]! : undefined;
  let value = cell ? getTableCellText(cell) : getShapeText(shape);
  for (const change of changes) {
    value = value.slice(0, change.start) + change.text + value.slice(change.end);
    if (cell) setTableCellText(cell, value, { preserveFormatting: true });
    else setShapeText(shape, value, { preserveFormatting: true });
  }
}

/** A disposable model for reading pending run formats without touching history. */
export function projectTextEdits(
  shape: SlideShapeData,
  changes: readonly TextEdit[],
  position?: CellPosition,
): SlideShapeData {
  if (!changes.length) return shape;
  const copy = copyShape(addBlankSlide(createPresentation()), shape);
  replayTextEdits(copy, changes, position);
  return copy;
}
