import {
  getShapeParagraphCount,
  getShapeRunFormatEffective,
  type PresentationData,
  getParagraphEndFormat,
  getShapeParagraphElements,
  getTableCells,
  getTableCellParagraphs,
  type SlideShapeData,
  type TextFormat,
  toWritableTextFormat,
} from '@office-kit/pptx';

/** Character formats at a caret or intersecting a selected text range. */
export function textFormatsInRange(
  shape: SlideShapeData,
  range: { start: number; end: number },
  cell?: { row: number; col: number },
  context?: { pres: PresentationData; source?: SlideShapeData },
): TextFormat[] {
  const formats: TextFormat[] = [];
  let offset = 0;
  const paragraphs = cell
    ? getTableCellParagraphs(getTableCells(shape)[cell.row]![cell.col]!)
    : Array.from({ length: getShapeParagraphCount(shape) }, (_, i) => ({
        elements: getShapeParagraphElements(shape, i),
        endFormat: getParagraphEndFormat(shape, i),
      }));
  for (const [paragraphIndex, { elements, endFormat }] of paragraphs.entries()) {
    let runIndex = 0;
    const paragraphStart = offset;
    for (const element of elements) {
      const currentRun = runIndex;
      if (element.kind === 'r') runIndex++;
      // Readers widen colors to plain strings; the selection's format is fed
      // straight back into writers, so it is converted once here.
      const format = (): TextFormat =>
        toWritableTextFormat(
          context && !cell && element.kind === 'r'
            ? getShapeRunFormatEffective(context.pres, shape, paragraphIndex, currentRun, {
                inheritanceSource: context.source ?? shape,
              })
            : (element.format ?? {}),
        );
      const length = element.kind === 'br' ? 1 : element.text.length;
      if (range.start === range.end) {
        const caret = range.start;
        if (
          length &&
          ((offset < caret && offset + length >= caret) ||
            (caret === paragraphStart && offset === caret))
        )
          return [format()];
      } else if (offset < range.end && offset + length > range.start) formats.push(format());
      offset += length;
    }
    // Empty paragraphs have no character to sample. Their end mark carries
    // the format inherited by the next typed character.
    if (offset === paragraphStart && range.start === offset && range.end === offset) {
      return [toWritableTextFormat(endFormat ?? elements[0]?.format ?? {})];
    }
    offset++;
  }
  return formats;
}
