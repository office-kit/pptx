import {
  addBlankSlide,
  copyShape,
  createPresentation,
  getShapeText,
  getShapeParagraphCount,
  getParagraphPropertiesEffective,
  setParagraphAlignment,
  setParagraphBullet,
  setParagraphLineSpacing,
  setParagraphSpacing,
  type PresentationData,
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

/** A disposable model for reading pending toolbar formats without touching history. */
export function projectTextEdits(
  shape: SlideShapeData,
  changes: readonly TextEdit[],
  position?: CellPosition,
  pres?: PresentationData,
): SlideShapeData {
  if (!changes.length) return shape;
  const copy = copyShape(addBlankSlide(createPresentation()), shape);
  // copyShape has no source layout/master. Preserve the effective paragraph
  // controls before edits split or merge paragraphs in the disposable shape.
  if (pres && !position) {
    for (let i = 0; i < getShapeParagraphCount(shape); i++) {
      const props = getParagraphPropertiesEffective(pres, shape, i);
      if (props.align !== null) setParagraphAlignment(copy, i, props.align);
      if (props.bullet !== null) setParagraphBullet(copy, i, props.bullet);
      if (props.lineSpacing !== null) setParagraphLineSpacing(copy, i, props.lineSpacing);
      setParagraphSpacing(copy, i, { beforePts: props.spcBefPts, afterPts: props.spcAftPts });
    }
  }
  replayTextEdits(copy, changes, position);
  return copy;
}
