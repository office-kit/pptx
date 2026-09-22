import { expect, it } from 'vitest';
import { textEditDiff } from '../site/src/lib/editor/core/text-edit-diff.ts';
import { projectTextEdits } from '../site/src/lib/editor/core/text-edit-preview.ts';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  inches,
  setShapeParagraphs,
  getShapeParagraphElements,
  getTableCells,
  getTableCellParagraphs,
  setTableCellTextFormat,
} from '../src/api/index.ts';

it.each([
  ['aaa', 'aaaa', 1, 1, 2, { start: 1, end: 1, text: 'a' }],
  ['aaa', 'aa', 2, 2, 1, { start: 1, end: 2, text: '' }],
  ['aaa', 'aa', 1, 1, 1, { start: 1, end: 2, text: '' }],
  ['aaa', 'aa', 0, 2, 1, { start: 0, end: 2, text: 'a' }],
  ['A😀B', 'A😁B', 3, 3, 3, { start: 1, end: 3, text: '😁' }],
  ['日本語', '日本語です', 3, 3, 5, { start: 3, end: 3, text: 'です' }],
] as const)(
  'preserves the edit location for %s → %s',
  (before, value, start, end, after, expected) => {
    expect(textEditDiff(before, value, { start, end }, after)).toEqual(expected);
  },
);

for (const kind of ['shape', 'cell'] as const)
  it(`keeps distinct formats on repeated characters in a ${kind}`, () => {
    const slide = addBlankSlide(createPresentation());
    const bounds = { x: inches(1), y: inches(1), w: inches(4), h: inches(2) };
    const shape =
      kind === 'shape'
        ? addSlideTextBox(slide, { ...bounds, text: '' })
        : addSlideTable(slide, { ...bounds, rows: [['aaa', 'Other']] });
    if (kind === 'shape')
      setShapeParagraphs(shape, [
        {
          runs: [
            { text: 'a', format: { bold: true } },
            { text: 'a', format: { italic: true } },
            { text: 'a', format: { underline: true } },
          ],
        },
      ]);
    else {
      const cell = getTableCells(shape)[0]![0]!;
      setTableCellTextFormat(cell, { bold: true }, { range: { start: 0, end: 1 } });
      setTableCellTextFormat(cell, { italic: true }, { range: { start: 1, end: 2 } });
      setTableCellTextFormat(cell, { underline: true }, { range: { start: 2, end: 3 } });
    }
    const position = kind === 'cell' ? { row: 0, col: 0 } : undefined;
    const formats = (edited: typeof shape) => {
      const elements =
        kind === 'shape'
          ? getShapeParagraphElements(edited, 0)
          : getTableCellParagraphs(getTableCells(edited)[0]![0]!)[0]!.elements;
      return elements.flatMap((r) => (r.kind === 'br' ? [] : Array.from(r.text, () => r.format)));
    };
    const original = formats(shape);
    expect(original).toMatchObject([{ bold: true }, { italic: true }, { underline: true }]);
    const inserted = projectTextEdits(shape, [{ start: 1, end: 1, text: 'a' }], position);
    expect(formats(inserted)).toEqual([original[0], original[0], original[1], original[2]]);
    const deleted = projectTextEdits(shape, [{ start: 1, end: 2, text: '' }], position);
    expect(formats(deleted)).toEqual([original[0], original[2]]);
    expect(formats(shape)).toEqual(original);
  });
