import {
  getShapeParagraphCount,
  getParagraphEndFormat,
  getShapeParagraphElements,
  getTableCells,
  getTableCellParagraphs,
  type SlideShapeData,
  type TextFormat,
} from '@office-kit/pptx';

/** Character formats at a caret or intersecting a selected text range. */
export function textFormatsInRange(
  shape: SlideShapeData,
  range: { start: number; end: number },
  cell?: { row: number; col: number },
): TextFormat[] {
  const formats: TextFormat[] = [];
  let offset = 0;
  const paragraphs = cell
    ? getTableCellParagraphs(getTableCells(shape)[cell.row]![cell.col]!)
    : Array.from({ length: getShapeParagraphCount(shape) }, (_, i) => ({
        elements: getShapeParagraphElements(shape, i),
        endFormat: getParagraphEndFormat(shape, i),
      }));
  for (const { elements, endFormat } of paragraphs) {
    const paragraphStart = offset;
    for (const element of elements) {
      const length = element.kind === 'br' ? 1 : element.text.length;
      if (range.start === range.end) {
        const caret = range.start;
        if (
          length &&
          ((offset < caret && offset + length >= caret) ||
            (caret === paragraphStart && offset === caret))
        )
          return [element.format ?? {}];
      } else if (offset < range.end && offset + length > range.start)
        formats.push(element.format ?? {});
      offset += length;
    }
    // Empty paragraphs have no character to sample. Their end mark carries
    // the format inherited by the next typed character.
    if (offset === paragraphStart && range.start === offset && range.end === offset) {
      return [endFormat ?? elements[0]?.format ?? {}];
    }
    offset++;
  }
  return formats;
}
