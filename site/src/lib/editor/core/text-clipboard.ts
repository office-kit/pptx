import {
  getShapeParagraphCount,
  getShapeParagraphElements,
  getTableCellParagraphs,
  getTableCells,
  type SlideShapeData,
  type TextFormat,
} from '@office-kit/pptx';
import type { TextEdit } from './text-edit-preview.ts';

export const TEXT_CLIPBOARD_TYPE = 'application/x-office-kit-text+json';
type TextClipboard = { version: 1; text: string; formats: NonNullable<TextEdit['formats']> };

export function copyTextRange(
  shape: SlideShapeData,
  start: number,
  end: number,
  cell?: { row: number; col: number },
): TextClipboard {
  const paragraphs = cell
    ? getTableCellParagraphs(getTableCells(shape)[cell.row]![cell.col]!).map((p) => p.elements)
    : Array.from({ length: getShapeParagraphCount(shape) }, (_, i) =>
        getShapeParagraphElements(shape, i),
      );
  let text = '';
  const formats: TextClipboard['formats'] = [];
  function append(value: string, format: TextFormat) {
    const offset = text.length;
    text += value;
    const from = Math.max(start, offset);
    const to = Math.min(end, text.length);
    if (from < to) formats.push({ start: from - start, end: to - start, format: { ...format } });
  }
  paragraphs.forEach((elements, index) => {
    if (index) append('\n', {});
    for (const element of elements)
      append(element.kind === 'br' ? '\n' : element.text, element.format ?? {});
  });
  return { version: 1, text: text.slice(start, end), formats };
}

function validFormat(value: unknown): value is TextFormat {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  return Object.entries(value).every(([key, v]) => {
    switch (key) {
      case 'font':
      case 'fontEastAsian':
      case 'fontComplexScript':
        return typeof v === 'string' && v.length <= 256;
      case 'color':
      case 'highlight':
        return (
          v === null ||
          (typeof v === 'string' &&
            /^(#?[\da-f]{6}|dk[12]|lt[12]|tx[12]|bg[12]|accent[1-6]|hlink|folHlink)$/i.test(v))
        );
      case 'bold':
      case 'italic':
        return typeof v === 'boolean';
      case 'underline':
      case 'strike':
        return typeof v === 'boolean' || (typeof v === 'string' && /^[a-zA-Z]{1,32}$/.test(v));
      case 'cap':
        return v === 'none' || v === 'small' || v === 'all';
      case 'size':
        return typeof v === 'number' && Number.isFinite(v) && v > 0 && v <= 4000;
      case 'spc':
      case 'kern':
      case 'baseline':
        return typeof v === 'number' && Number.isFinite(v) && Math.abs(v) <= 1_000_000;
      default:
        return false;
    }
  });
}

/** Foreign or malformed clipboard metadata falls back to the native plain-text paste. */
export function parseTextClipboard(raw: string, plain: string): TextClipboard | null {
  const maxClipboardLength = 4_000_000;
  if (!raw || raw.length > maxClipboardLength) return null;
  let value;
  try {
    value = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!value || value.version !== 1 || value.text !== plain || !Array.isArray(value.formats))
    return null;
  let offset = 0;
  for (const span of value.formats) {
    if (
      !span ||
      !Number.isInteger(span.start) ||
      !Number.isInteger(span.end) ||
      span.start !== offset ||
      span.end <= span.start ||
      span.end > plain.length ||
      !validFormat(span.format)
    )
      return null;
    offset = span.end;
  }
  if (offset !== plain.length) return null;
  return { version: 1, text: plain, formats: value.formats };
}
