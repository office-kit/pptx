import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  groupShapes,
  inches,
  setShapeHidden,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

it('omits hidden objects and hidden group descendants without deleting their content', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shapes = ['Visible label', 'Hidden label', 'Grouped label', 'Sibling label'].map(
    (text, i) =>
      addSlideTextBox(slide, { x: inches(i), y: inches(1), w: inches(2), h: inches(1), text }),
  );
  const group = groupShapes([shapes[2]!, shapes[3]!]);
  setShapeHidden(shapes[1]!, true);
  setShapeHidden(group, true);
  const svg = renderSlideToSvg(pres, slide);
  expect(svg).toContain('Visible label');
  expect(svg).not.toContain('Hidden label');
  expect(svg).not.toContain('Grouped label');
  setShapeHidden(group, false);
  expect(renderSlideToSvg(pres, slide)).toContain('Grouped label');
});
