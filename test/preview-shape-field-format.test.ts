import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  getShapeRunFormatEffective,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { textFormatsInRange } from '../site/src/lib/editor/core/text-format-selection.ts';

it.each(['svg', 'foreignObject'] as const)(
  'shape fields inherit paragraph formatting (%s)',
  async (textLayout) => {
    const original = createPresentation();
    addSlideTextBox(addBlankSlide(original), {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'Date',
    });
    const { entries } = readZip(await savePresentation(original));
    const load = (field: boolean) =>
      loadPresentation(
        writeZip(
          entries.map((entry) => {
            if (entry.name !== 'ppt/slides/slide1.xml') return entry;
            let xml = new TextDecoder()
              .decode(entry.data)
              .replace(
                /<a:p>[\s\S]*?<\/a:p>/,
                '<a:p><a:pPr><a:defRPr sz="2800" b="1"><a:solidFill><a:srgbClr val="AA2244"/></a:solidFill><a:latin typeface="Courier New"/></a:defRPr></a:pPr><a:r><a:rPr i="1"/><a:t>Date</a:t></a:r></a:p>',
              );
            if (field)
              xml = xml
                .replace(
                  '<a:r>',
                  '<a:fld id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" type="datetime1">',
                )
                .replace('</a:r>', '</a:fld>');
            return { ...entry, data: new TextEncoder().encode(xml) };
          }),
        ),
      );
    const regular = await load(false);
    const field = await load(true);
    expect(renderSlideToSvg(field, getSlides(field)[0]!, { textLayout })).toBe(
      renderSlideToSvg(regular, getSlides(regular)[0]!, { textLayout }),
    );
    const shape = getSlideShapes(getSlides(field)[0]!)[0]!;
    const expected = { size: 28, bold: true, italic: true, font: 'Courier New', color: '#AA2244' };
    expect(textFormatsInRange(shape, { start: 0, end: 4 }, undefined, { pres: field })).toEqual([
      expect.objectContaining(expected),
    ]);
    expect(getShapeRunFormatEffective(field, shape, 0, { fieldIndex: 0 })).toMatchObject(expected);
    expect(() => getShapeRunFormatEffective(field, shape, 0, 0)).toThrow(RangeError);
  },
);
