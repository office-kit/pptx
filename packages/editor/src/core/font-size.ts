import {
  getShapeParagraphCount,
  getShapeParagraphElements,
  getShapeText,
  getParagraphEndFormat,
  setShapeTextFormat,
  getTableCellParagraphs,
  setTableCellTextFormat,
  toWritableTextFormat,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { textFormatsInRange } from './text-format-selection.ts';
import { DEFAULT_BODY_PT, defaultTextMetrics } from './text-layout-defaults.ts';

/**
 * The sizes in PowerPoint's font-size gallery.  Grow/Shrink follows this
 * ladder while the value is in the gallery; values above the gallery use the
 * same 20% scaling PowerPoint uses for custom sizes.
 */
export const FONT_SIZE_GALLERY = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
] as const;

const GALLERY_MAX = FONT_SIZE_GALLERY[FONT_SIZE_GALLERY.length - 1];

/** Return the size PowerPoint selects for Home > Grow/Shrink Font. */
export function stepFontSize(size: number, direction: 1 | -1): number {
  if (size <= 0) return size;
  // PowerPoint walks integer sizes below the gallery's 8pt floor.
  if (size < 8)
    return direction > 0 ? Math.min(8, Math.floor(size) + 1) : Math.max(1, Math.ceil(size) - 1);
  if (size > GALLERY_MAX) {
    return Math.min(
      4000,
      direction > 0 ? Math.round(size * 1.2) : Math.max(GALLERY_MAX, Math.round(size / 1.2)),
    );
  }
  if (direction > 0) {
    return Math.min(
      4000,
      FONT_SIZE_GALLERY.find((candidate) => candidate > size) ?? Math.round(size * 1.2),
    );
  }
  for (let index = FONT_SIZE_GALLERY.length - 1; index >= 0; index--) {
    const candidate = FONT_SIZE_GALLERY[index]!;
    if (candidate < size) return candidate;
  }
  return Math.max(1, Math.ceil(size) - 1);
}

/** Apply a relative size to every visible run, retaining mixed run sizes. */
export function stepShapeFontSize(
  pres: PresentationData,
  shape: SlideShapeData,
  direction: 1 | -1,
  source = shape,
  range?: { start: number; end: number },
): boolean {
  let offset = 0;
  const total = getShapeText(shape).length;
  const defaultSize = defaultTextMetrics(pres, source).size;
  const formats = textFormatsInRange(shape, { start: 0, end: total }, undefined, { pres, source });
  let formatIndex = 0;
  let changed = false;
  for (let paragraphIndex = 0; paragraphIndex < getShapeParagraphCount(shape); paragraphIndex++) {
    const elements = getShapeParagraphElements(shape, paragraphIndex);
    if (!elements.some((element) => (element.kind === 'br' ? 1 : element.text.length) > 0)) {
      const selected =
        !range ||
        (range.start === range.end
          ? range.start === offset
          : range.start <= offset && range.end > offset);
      if (selected) {
        const endFormat = getParagraphEndFormat(shape, paragraphIndex);
        const emptyRun = elements.find((element) => element.kind !== 'br');
        const currentSize =
          endFormat?.size ??
          (emptyRun?.kind === 'r' ? emptyRun.format?.size : undefined) ??
          defaultSize;
        setShapeTextFormat(
          shape,
          { size: stepFontSize(currentSize, direction) },
          { paragraphEnd: paragraphIndex },
        );
        changed = true;
      }
    }
    for (const element of elements) {
      const length = element.kind === 'br' ? 1 : element.text.length;
      const elementStart = offset;
      const elementEnd = offset + length;
      const format = length > 0 ? formats[formatIndex++] : undefined;
      if (length > 0 && (!range || (elementStart < range.end && elementEnd > range.start))) {
        const selectedStart = range ? Math.max(elementStart, range.start) : elementStart;
        const selectedEnd = range ? Math.min(elementEnd, range.end) : elementEnd;
        if (format) {
          setShapeTextFormat(
            shape,
            { size: stepFontSize(format.size ?? defaultSize, direction) },
            { range: { start: selectedStart, end: selectedEnd } },
          );
          changed = true;
        }
      }
      offset = elementEnd;
    }
    offset++;
  }
  return changed;
}

export function stepTableCellFontSize(
  cell: Parameters<typeof setTableCellTextFormat>[0],
  direction: 1 | -1,
  range?: { start: number; end: number },
): boolean {
  let offset = 0;
  let changed = false;
  for (const [paragraphIndex, paragraph] of getTableCellParagraphs(cell).entries()) {
    const elements = paragraph.elements;
    if (!elements.some((element) => (element.kind === 'br' ? 1 : element.text.length) > 0)) {
      const selected =
        !range ||
        (range.start === range.end
          ? range.start === offset
          : range.start <= offset && range.end > offset);
      if (selected) {
        const emptyRun = elements.find((element) => element.kind !== 'br');
        const currentSize =
          paragraph.endFormat?.size ?? (emptyRun?.kind === 'r' ? emptyRun.format?.size : undefined);
        setTableCellTextFormat(
          cell,
          { size: stepFontSize(currentSize ?? DEFAULT_BODY_PT, direction) },
          { paragraphEnd: paragraphIndex },
        );
        changed = true;
      }
    }
    for (const element of elements) {
      const length = element.kind === 'br' ? 1 : element.text.length;
      const start = offset;
      const end = offset + length;
      if (length > 0 && (!range || (start < range.end && end > range.start))) {
        const selectedStart = range ? Math.max(start, range.start) : start;
        const selectedEnd = range ? Math.min(end, range.end) : end;
        const format = toWritableTextFormat(element.format ?? {});
        setTableCellTextFormat(
          cell,
          { size: stepFontSize(format.size ?? DEFAULT_BODY_PT, direction) },
          { range: { start: selectedStart, end: selectedEnd } },
        );
        changed = true;
      }
      offset = end;
    }
    offset++;
  }
  return changed;
}
