import { expect, it } from 'vitest';
import {
  createPresentation,
  inches,
  addBlankSlide,
  addSlideTextBox,
  getShapeText,
  getSlides,
  getSlideShapes,
  getShapeRunHyperlink,
  getParagraphLevel,
  getShapeParagraphElements,
  setShapeParagraphs,
  setShapeRunHyperlink,
  setParagraphLevel,
  savePresentation,
  loadPresentation,
} from '../src/api/index.ts';

it('concatenates paragraphs across slides without losing formatting, links or source text', async () => {
  const pres = createPresentation();
  const first = addBlankSlide(pres);
  const second = addBlankSlide(pres);
  const box = (slide: typeof first, text: string) =>
    addSlideTextBox(slide, { x: inches(0), y: inches(0), w: inches(1), h: inches(1), text });
  const target = box(first, 'Existing');
  const heading = box(second, 'Heading');
  const body = box(second, 'Child\n');
  setShapeParagraphs(heading, [{ runs: [{ text: 'Heading', format: { bold: true } }] }]);
  setShapeRunHyperlink(target, 0, 0, 'https://example.com/existing');
  setShapeRunHyperlink(heading, 0, 0, 'https://example.com/heading');
  setShapeRunHyperlink(body, 0, 0, 'https://example.com/child');
  setParagraphLevel(body, 0, 2);
  setShapeParagraphs(target, { sources: [target, heading, body] });
  expect(getShapeText(heading)).toBe('Heading');
  expect(getShapeText(body)).toBe('Child\n');
  const loaded = await loadPresentation(await savePresentation(pres));
  const result = getSlideShapes(getSlides(loaded)[0]!)[0]!;
  expect(getShapeText(result)).toBe('Existing\nHeading\nChild\n');
  expect(getShapeParagraphElements(result, 1)[0]?.format?.bold).toBe(true);
  expect(getParagraphLevel(result, 2)).toBe(2);
  expect([0, 1, 2].map((index) => getShapeRunHyperlink(result, index, 0))).toEqual([
    'https://example.com/existing',
    'https://example.com/heading',
    'https://example.com/child',
  ]);
});

it('rejects empty or sparse sources before replacing the target', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const target = addSlideTextBox(slide, {
    x: inches(0),
    y: inches(0),
    w: inches(1),
    h: inches(1),
    text: 'Kept',
  });
  expect(() => setShapeParagraphs(target, { sources: [] })).toThrow(RangeError);
  const sources = [target];
  sources.length = 2;
  expect(() => setShapeParagraphs(target, { sources })).toThrow(TypeError);
  expect(getShapeText(target)).toBe('Kept');
});

it('concatenates ranges of a source into itself without losing levels, formats or links', async () => {
  const pres = createPresentation();
  const source = addSlideTextBox(addBlankSlide(pres), {
    x: inches(0),
    y: inches(0),
    w: inches(2),
    h: inches(2),
    text: '',
  });
  setShapeParagraphs(source, [
    { runs: [{ text: 'First' }] },
    { runs: [{ text: 'Second', format: { italic: true } }] },
    { runs: [{ text: 'Third' }] },
    { runs: [] },
  ]);
  setShapeRunHyperlink(source, 1, 0, 'https://example.com/second');
  setParagraphLevel(source, 2, 2);
  setShapeParagraphs(source, {
    source,
    ranges: [
      { start: 6, end: 12 },
      { start: 0, end: 5 },
      { start: 13, end: 19 },
    ],
  });
  const loaded = await loadPresentation(await savePresentation(pres));
  const result = getSlideShapes(getSlides(loaded)[0]!)[0]!;
  expect(getShapeText(result)).toBe('Second\nFirst\nThird\n');
  expect(getShapeParagraphElements(result, 0)[0]?.format?.italic).toBe(true);
  expect(getShapeRunHyperlink(result, 0, 0)).toBe('https://example.com/second');
  expect(getParagraphLevel(result, 2)).toBe(2);
  expect(() => setShapeParagraphs(source, { source, ranges: [] })).toThrow(RangeError);
  expect(() => setShapeParagraphs(source, { source, ranges: [{ start: 0, end: 50 }] })).toThrow();
  expect(getShapeText(source)).toBe('Second\nFirst\nThird\n');
});
