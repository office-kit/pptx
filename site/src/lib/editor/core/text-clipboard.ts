import {
  asColor,
  getShapeParagraphElements,
  getTableCellParagraphs,
  getTableCells,
  type SlideShapeData,
  type TextFormat,
  toWritableTextFormat,
} from '@office-kit/pptx';
import type { TextEdit } from './text-edit-preview.ts';

export const TEXT_CLIPBOARD_TYPE = 'application/x-office-kit-text+json';
type TextClipboard = { version: 1; text: string; formats: NonNullable<TextEdit['formats']> };

// DrawingML ST_TextUnderlineType values accepted by the text writer.
const underlineStyles = new Set([
  'none',
  'words',
  'sng',
  'dbl',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDbl',
]);
const strikeStyles = new Set(['noStrike', 'sngStrike', 'dblStrike']);
const percentageScale = 100_000;
const minPercentageInteger = -2_147_483_648;
const maxPercentageInteger = 2_147_483_647;

export function copyTextRange(
  shape: SlideShapeData,
  start: number,
  end: number,
  cell?: { row: number; col: number },
  resolveRunFormat?: (paragraphIndex: number, runIndex: number) => TextFormat,
): TextClipboard {
  const paragraphs = cell
    ? getTableCellParagraphs(getTableCells(shape)[cell.row]![cell.col]!).map((p) => p.elements)
    : getShapeParagraphElements(shape);
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
    let runIndex = 0;
    for (const element of elements) {
      let format = element.format;
      if (element.kind === 'r') {
        if (resolveRunFormat && text.length < end && text.length + element.text.length > start)
          format = resolveRunFormat(index, runIndex);
        runIndex++;
      }
      // What the reader hands back widens colors to strings; the clipboard
      // carries a format that can be written straight into another shape.
      append(element.kind === 'br' ? '\n' : element.text, toWritableTextFormat(format ?? {}));
    }
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
        return v === null || (typeof v === 'string' && asColor(v) !== null);
      case 'bold':
      case 'italic':
      case 'normalizeHeight':
        return typeof v === 'boolean';
      case 'underline':
        return typeof v === 'boolean' || (typeof v === 'string' && underlineStyles.has(v));
      case 'strike':
        return typeof v === 'boolean' || (typeof v === 'string' && strikeStyles.has(v));
      case 'cap':
        return v === 'none' || v === 'small' || v === 'all';
      case 'size':
        return typeof v === 'number' && Number.isFinite(v) && v >= 1 && v <= 4000;
      case 'spc':
        return typeof v === 'number' && Number.isFinite(v) && Math.abs(v) <= 400_000;
      case 'kern':
        return typeof v === 'number' && Number.isFinite(v) && v >= 0 && v <= 400_000;
      case 'baseline':
        // The writer emits ST_PercentageDecimal as a signed 32-bit integer.
        return (
          typeof v === 'number' &&
          Number.isFinite(v) &&
          Math.round(v * percentageScale) >= minPercentageInteger &&
          Math.round(v * percentageScale) <= maxPercentageInteger
        );
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
