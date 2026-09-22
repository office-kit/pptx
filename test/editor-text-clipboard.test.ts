import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  inches,
  setShapeParagraphs,
  getShapeParagraphElements,
  getShapeText,
  getTableCells,
  getTableCellParagraphs,
} from '../src/api/index.ts';
import { copyTextRange, parseTextClipboard } from '../site/src/lib/editor/core/text-clipboard.ts';
import { projectTextEdits } from '../site/src/lib/editor/core/text-edit-preview.ts';

function fixture() {
  const slide = addBlankSlide(createPresentation());
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    text: '',
  });
  setShapeParagraphs(shape, [
    {
      runs: [
        { text: 'English', format: { bold: true, color: '#13579B', size: 32 } },
        { text: '日本語', format: { italic: true } },
      ],
    },
  ]);
  return { slide, shape };
}

describe('formatted text clipboard', () => {
  it('copies clipped mixed-format runs including pending typing, without mutating the source', () => {
    const { shape } = fixture();
    const pending = projectTextEdits(shape, [
      { start: 10, end: 10, text: 'です', typing: { format: { underline: true }, reset: true } },
    ]);
    const copied = copyTextRange(pending, 4, 12);
    expect(copied.text).toBe('ish日本語です');
    expect(copied.formats.map((s) => [s.start, s.end])).toEqual([
      [0, 3],
      [3, 6],
      [6, 8],
    ]);
    expect(copied.formats[0]!.format).toMatchObject({ bold: true, size: 32, color: '#13579B' });
    expect(copied.formats[1]!.format.italic).toBe(true);
    expect(parseTextClipboard(JSON.stringify(copied), copied.text)).toEqual(copied);
    expect(getShapeText(shape)).toBe('English日本語');
  });
  it('replaces selected text with separate formats and preserves surrounding characters', () => {
    const { shape } = fixture();
    const copied = copyTextRange(shape, 5, 10);
    const result = projectTextEdits(shape, [
      { start: 1, end: 2, text: copied.text, formats: copied.formats },
    ]);
    expect(getShapeText(result)).toBe('Esh日本語glish日本語');
    const runs = getShapeParagraphElements(result, 0);
    const chars = runs.flatMap((r) =>
      r.kind === 'br' ? [] : Array.from(r.text, (c) => ({ c, format: r.format })),
    );
    expect(chars[3]!.format?.italic).toBe(true);
    expect(chars[3]!.format?.bold).not.toBe(true);
    expect(chars[6]!.format?.bold).toBe(true);
  });
  it('round-trips paragraph breaks and copies shape runs into a table cell', () => {
    const { slide, shape } = fixture();
    setShapeParagraphs(shape, [
      { runs: [{ text: 'Hi', format: { bold: true } }] },
      { runs: [{ text: '日本語', format: { italic: true } }] },
    ]);
    const copied = copyTextRange(shape, 1, 6);
    expect(copied.text).toBe('i\n日本語');
    expect(parseTextClipboard(JSON.stringify(copied), copied.text)).toEqual(copied);
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      rows: [['old']],
    });
    const result = projectTextEdits(
      table,
      [{ start: 0, end: 3, text: copied.text, formats: copied.formats }],
      { row: 0, col: 0 },
    );
    const paragraphs = getTableCellParagraphs(getTableCells(result)[0]![0]!);
    expect(paragraphs[0]!.elements[0]!.format?.bold).toBe(true);
    expect(paragraphs[1]!.elements[0]!.format?.italic).toBe(true);
    expect(copyTextRange(result, 0, 5, { row: 0, col: 0 }).text).toBe(copied.text);
  });
  it('rejects malformed, mismatched and unsafe metadata for plain text fallback', () => {
    const { shape } = fixture();
    const copied = copyTextRange(shape, 0, 7);
    expect(parseTextClipboard('{', copied.text)).toBeNull();
    expect(parseTextClipboard(JSON.stringify(copied), 'different')).toBeNull();
    for (const formats of [
      [],
      [{ start: 1, end: 7, format: {} }],
      [{ start: 0, end: 99, format: {} }],
      [{ start: 0, end: 7, format: { size: -2 } }],
      [{ start: 0, end: 7, format: { evil: true } }],
    ]) {
      expect(parseTextClipboard(JSON.stringify({ ...copied, formats }), copied.text)).toBeNull();
    }
  });
});
