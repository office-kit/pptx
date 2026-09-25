import {
  addBlankSlide,
  copyShape,
  createPresentation,
  getShapeParagraphCount,
  getParagraphPropertiesEffective,
  setParagraphAlignment,
  setParagraphBullet,
  setParagraphLineSpacing,
  setParagraphSpacing,
  type PresentationData,
  getTableCells,
  setShapeText,
  setShapeTextFormat,
  setTableCellTextFormat,
  type TextFormat,
  setTableCellText,
  type SlideShapeData,
} from '@office-kit/pptx';

export type TextEdit = {
  start: number;
  end: number;
  text: string;
  typing?: { format: TextFormat; reset: boolean };
  formats?: { start: number; end: number; format: TextFormat }[];
};
type CellPosition = { row: number; col: number };

/** Use the same formatting-preserving edits for preview and commit. */
export function replayTextEdits(
  shape: SlideShapeData,
  changes: readonly TextEdit[],
  position?: CellPosition,
): void {
  const cell = position ? getTableCells(shape)[position.row]![position.col]! : undefined;
  for (const change of changes) {
    const range = { start: change.start, end: change.end };
    if (cell) setTableCellText(cell, change.text, { range });
    else setShapeText(shape, change.text, { range });
    for (const span of change.formats ?? []) {
      const options = {
        range: { start: change.start + span.start, end: change.start + span.end },
        reset: true,
      };
      if (cell) setTableCellTextFormat(cell, span.format, options);
      else setShapeTextFormat(shape, span.format, options);
    }
    if (change.typing && change.text.length) {
      const options = {
        range: { start: change.start, end: change.start + change.text.length },
        reset: change.typing.reset,
      };
      if (cell) setTableCellTextFormat(cell, change.typing.format, options);
      else setShapeTextFormat(shape, change.typing.format, options);
    }
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
