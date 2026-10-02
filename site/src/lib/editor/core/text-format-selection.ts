import {
  getShapeParagraphCount,
  getShapeRunFormatEffective,
  type PresentationData,
  getParagraphEndFormat,
  getShapeParagraphElements,
  getTableCells,
  getTableCellParagraphs,
  getTableCellRunFormatEffective,
  type SlideShapeData,
  type TextFormat,
  toWritableTextFormat,
} from '@office-kit/pptx';
import { defaultTextMetrics } from './text-layout-defaults.ts';

/** Character formats at a caret or intersecting a selected text range. */
export function textFormatsInRange(
  shape: SlideShapeData,
  range: { start: number; end: number },
  cell?: { row: number; col: number },
  context?: { pres: PresentationData; source?: SlideShapeData },
): TextFormat[] {
  const formats: TextFormat[] = [];
  const defaultSize = context
    ? defaultTextMetrics(context.pres, context.source ?? shape).size
    : undefined;
  function selectionFormat(format: Parameters<typeof toWritableTextFormat>[0]): TextFormat {
    const writable = toWritableTextFormat(format);
    // The controls must show the size painted by the editor even when XML omits it.
    return defaultSize === undefined
      ? writable
      : { ...writable, size: writable.size ?? defaultSize };
  }
  let offset = 0;
  const tableCell = cell ? getTableCells(shape)[cell.row]![cell.col]! : undefined;
  const paragraphs = tableCell
    ? getTableCellParagraphs(tableCell)
    : Array.from({ length: getShapeParagraphCount(shape) }, (_, i) => ({
        elements: getShapeParagraphElements(shape, i),
        endFormat: getParagraphEndFormat(shape, i),
      }));
  for (const [paragraphIndex, { elements, endFormat }] of paragraphs.entries()) {
    let runIndex = 0;
    let fieldIndex = 0;
    let breakIndex = 0;
    const paragraphStart = offset;
    for (const element of elements) {
      const currentRun = runIndex;
      const currentField = fieldIndex;
      const currentBreak = breakIndex;
      if (element.kind === 'br') breakIndex++;
      if (element.kind === 'fld') fieldIndex++;
      if (element.kind === 'r') runIndex++;
      // Readers widen colors to plain strings; the selection's format is fed
      // straight back into writers, so it is converted once here.
      const format = (): TextFormat =>
        selectionFormat(
          context
            ? tableCell
              ? getTableCellRunFormatEffective(
                  context.pres,
                  tableCell,
                  paragraphIndex,
                  element.kind === 'fld'
                    ? { fieldIndex: currentField }
                    : element.kind === 'br'
                      ? { breakIndex: currentBreak }
                      : currentRun,
                )
              : getShapeRunFormatEffective(
                  context.pres,
                  shape,
                  paragraphIndex,
                  element.kind === 'fld'
                    ? { fieldIndex: currentField }
                    : element.kind === 'br'
                      ? { breakIndex: currentBreak }
                      : currentRun,
                  { inheritanceSource: context.source ?? shape },
                )
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
      } else if (length > 0 && offset < range.end && offset + length > range.start)
        formats.push(format());
      offset += length;
    }
    // Empty paragraphs have no character to sample. Their end mark carries
    // the format inherited by the next typed character.
    if (offset === paragraphStart && range.start === offset && range.end === offset) {
      return [
        selectionFormat(
          context && elements.length === 0
            ? tableCell
              ? getTableCellRunFormatEffective(context.pres, tableCell, paragraphIndex, null)
              : getShapeRunFormatEffective(context.pres, shape, paragraphIndex, null, {
                  inheritanceSource: context.source ?? shape,
                })
            : (endFormat ?? elements[0]?.format ?? {}),
        ),
      ];
    }
    offset++;
  }
  return formats;
}
