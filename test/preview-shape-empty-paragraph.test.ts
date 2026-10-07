import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  inches,
  setShapeParagraphs,
  getSlides,
  getSlideShapes,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { textFormatsInRange } from '../packages/editor/src/core/text-format-selection.ts';

it.each(['svg', 'foreignObject'] as const)(
  'sizes empty shape paragraphs from their end mark (%s)',
  async (textLayout) => {
    const make = async (size: number) => {
      const pres = createPresentation();
      const box = addSlideTextBox(addBlankSlide(pres), {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        text: '',
      });
      setShapeParagraphs(box, [
        { runs: [], endFormat: { size, font: 'Courier New', bold: true } },
        { runs: [{ text: 'After', format: { size: 18 } }] },
      ]);
      return loadPresentation(await savePresentation(pres));
    };
    const small = await make(18);
    const large = await make(36);
    expect(renderSlideToSvg(large, getSlides(large)[0]!, { textLayout })).not.toBe(
      renderSlideToSvg(small, getSlides(small)[0]!, { textLayout }),
    );
    expect(
      textFormatsInRange(
        getSlideShapes(getSlides(large)[0]!)[0]!,
        { start: 0, end: 0 },
        undefined,
        { pres: large },
      ),
    ).toEqual([expect.objectContaining({ size: 36, font: 'Courier New', bold: true })]);
  },
);
