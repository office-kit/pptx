// `data-pptx-shape-id` / `data-pptx-paragraph` / `data-pptx-cell` let a
// preview map a click on the SVG back to the shape, paragraph and cell.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  findSlideLayout,
  getShapeId,
  getSlideShapes,
  groupShapes,
  inches,
  loadPresentation,
  setShapeParagraphs,
} from '../src/api/index.ts';
import { renderSlideToSvg, type TextLayoutMode } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));
const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  return { pres, slide };
};
const at = (x: number) => ({ x: inches(x), y: inches(1), w: inches(2), h: inches(1) });
const values = (svg: string, attribute: string) =>
  [...svg.matchAll(new RegExp(`${attribute}="([^"]*)"`, 'g'))].map((match) => match[1]);

describe('preview source attributes', () => {
  it('tags every shape, group members included, with its shape id', async () => {
    const { pres, slide } = await blankSlide();
    const first = addSlideTextBox(slide, { ...at(1), name: 'Label', text: 'A' });
    const second = addSlideTextBox(slide, { ...at(4), name: 'Label', text: 'B' });
    const a = addSlideShape(slide, { ...at(1), preset: 'rect' });
    const b = addSlideShape(slide, { ...at(4), preset: 'ellipse' });
    const [aId, bId] = [getShapeId(a), getShapeId(b)];
    const group = groupShapes([a, b]);
    const svg = renderSlideToSvg(pres, slide);
    const ids = values(svg, 'data-pptx-shape-id');
    // Flattened document order: each group is followed by its members.
    expect(ids).toEqual(getSlideShapes(slide).map((shape) => String(getShapeId(shape))));
    expect(ids).toEqual(
      expect.arrayContaining(
        [getShapeId(first), getShapeId(second), getShapeId(group), aId, bId].map(String),
      ),
    );
    expect(new Set(ids).size).toBe(ids.length);
  });

  for (const mode of ['foreignObject', 'svg'] as TextLayoutMode[]) {
    it(`numbers paragraphs and table cells (${mode})`, async () => {
      const { pres, slide } = await blankSlide();
      const text = addSlideTextBox(slide, { ...at(1), text: '' });
      setShapeParagraphs(
        text,
        ['one', 'two', 'three'].map((value) => ({ runs: [{ text: value }] })),
      );
      addSlideTable(slide, {
        ...at(4),
        rows: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      });
      const svg = renderSlideToSvg(pres, slide, { textLayout: mode });
      expect(values(svg, 'data-pptx-cell')).toEqual(['0,0', '0,1', '1,0', '1,1']);
      expect(values(svg, 'data-pptx-paragraph')).toEqual(['0', '1', '2', '0', '0', '0', '0']);
      expect(svg).toMatch(/data-pptx-cell="1,1">(?:(?!data-pptx-cell).)*d</);
    });
  }
});
