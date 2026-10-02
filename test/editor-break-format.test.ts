import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  toWritableTextFormat,
  addSlideTable,
  getTableCells,
  getShapeRunFormatEffective,
  getTableCellRunFormatEffective,
  createPresentation,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { copyTextRange } from '../site/src/lib/editor/core/text-clipboard.ts';
import { textFormatsInRange } from '../site/src/lib/editor/core/text-format-selection.ts';

it.each(['shape', 'table'] as const)(
  'line breaks inherit paragraph formatting in %s',
  async (kind) => {
    const original = createPresentation();
    const slide = addBlankSlide(original);
    const bounds = { x: inches(1), y: inches(1), w: inches(4), h: inches(2) };
    if (kind === 'shape') addSlideTextBox(slide, { ...bounds, text: 'Text' });
    else addSlideTable(slide, { ...bounds, rows: [['Text']] });
    const { entries } = readZip(await savePresentation(original));
    const pres = await loadPresentation(
      writeZip(
        entries.map((entry) =>
          entry.name !== 'ppt/slides/slide1.xml'
            ? entry
            : {
                ...entry,
                data: new TextEncoder().encode(
                  new TextDecoder()
                    .decode(entry.data)
                    .replace(
                      /<a:p>[\s\S]*?<\/a:p>/,
                      '<a:p><a:pPr><a:defRPr sz="2800" b="1"><a:latin typeface="Courier New"/></a:defRPr></a:pPr><a:br><a:rPr i="1"/></a:br></a:p>',
                    ),
                ),
              },
        ),
      ),
    );
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    expect(
      textFormatsInRange(
        shape,
        { start: 0, end: 1 },
        kind === 'table' ? { row: 0, col: 0 } : undefined,
        { pres },
      ),
    ).toEqual([
      expect.objectContaining({ size: 28, bold: true, italic: true, font: 'Courier New' }),
    ]);
    const cell = kind === 'table' ? getTableCells(shape)[0]![0]! : undefined;
    const resolve = (selector: number | { breakIndex: number }) =>
      cell
        ? getTableCellRunFormatEffective(pres, cell, 0, selector)
        : getShapeRunFormatEffective(pres, shape, 0, selector);
    expect(resolve({ breakIndex: 0 })).toMatchObject({
      size: 28,
      bold: true,
      italic: true,
      font: 'Courier New',
    });
    expect(() => resolve(0)).toThrow(RangeError);
    expect(() => resolve({ breakIndex: 1 })).toThrow(RangeError);
    const copied = copyTextRange(
      shape,
      0,
      1,
      cell ? { row: 0, col: 0 } : undefined,
      (paragraph, selector) =>
        toWritableTextFormat(
          cell
            ? getTableCellRunFormatEffective(pres, cell, paragraph, selector)
            : getShapeRunFormatEffective(pres, shape, paragraph, selector),
        ),
    );
    expect(copied.text).toBe('\n');
    expect(copied.formats[0]?.format).toMatchObject({
      size: 28,
      bold: true,
      italic: true,
      font: 'Courier New',
    });
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredShape = getSlideShapes(getSlides(restored)[0]!)[0]!;
    expect(
      kind === 'shape'
        ? getShapeRunFormatEffective(restored, restoredShape, 0, { breakIndex: 0 })
        : getTableCellRunFormatEffective(restored, getTableCells(restoredShape)[0]![0]!, 0, {
            breakIndex: 0,
          }),
    ).toEqual(resolve({ breakIndex: 0 }));
  },
);
