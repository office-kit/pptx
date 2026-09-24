import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  inches,
  bringShapeForward,
  sendShapeBackward,
  bringShapeToFront,
  sendShapeToBack,
  getSlideShapes,
  getShapeText,
  groupShapes,
  getGroupChildren,
  savePresentation,
  loadPresentation,
  getSlides,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { parseXml, NS } from '../src/internal/xml/index.ts';

function fixture() {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shapes = ['A', 'B', 'C', 'D', 'E'].map((text) =>
    addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(1), h: inches(1), text }),
  );
  return { pres, slide, shapes };
}

for (const [operation, expected] of [
  [bringShapeToFront, ['B', 'D', 'E', 'A', 'C']],
  [sendShapeToBack, ['A', 'C', 'B', 'D', 'E']],
  [bringShapeForward, ['B', 'A', 'D', 'C', 'E']],
  [sendShapeBackward, ['A', 'C', 'B', 'D', 'E']],
] as const) {
  it(`${operation.name} preserves selected stacking order, including after save/load`, async () => {
    const { pres, slide, shapes } = fixture();
    operation([shapes[2]!, shapes[0]!]);
    expect(getSlideShapes(slide).map(getShapeText)).toEqual(expected);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(loaded)[0]!).map(getShapeText)).toEqual(expected);
  });
}

it('moves contiguous selections together, clamps at the ends and accepts duplicates', () => {
  const { slide, shapes } = fixture();
  bringShapeForward([shapes[1]!, shapes[0]!, shapes[0]!]);
  expect(getSlideShapes(slide).map(getShapeText)).toEqual(['C', 'A', 'B', 'D', 'E']);
  sendShapeBackward([shapes[0]!, shapes[1]!]);
  expect(getSlideShapes(slide).map(getShapeText)).toEqual(['A', 'B', 'C', 'D', 'E']);
  sendShapeBackward(shapes.slice(0, 2));
  bringShapeForward(shapes.slice(3));
  expect(getSlideShapes(slide).map(getShapeText)).toEqual(['A', 'B', 'C', 'D', 'E']);
  bringShapeToFront([]);
});

it('reorders group siblings without moving extensions or accepting mixed parents', () => {
  const { slide, shapes } = fixture();
  const group = groupShapes(shapes.slice(0, 3));
  const extension = parseXml(`<p:extLst xmlns:p="${NS.pml}"><p:ext uri="keep"/></p:extLst>`).root;
  group[SHAPE_ELEMENT].children.push(extension);
  bringShapeToFront([shapes[1]!, shapes[0]!]);
  expect(getGroupChildren(group).map(getShapeText)).toEqual(['C', 'A', 'B']);
  expect(group[SHAPE_ELEMENT].children.at(-1)).toBe(extension);
  const before = getSlideShapes(slide).map(getShapeText);
  expect(() => sendShapeToBack([shapes[0]!, shapes[3]!])).toThrow(/same parent/);
  expect(getSlideShapes(slide).map(getShapeText)).toEqual(before);
});

it('rejects selections across slides without partially moving them', () => {
  const { shapes, slide, pres } = fixture();
  const other = addSlideTextBox(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(1),
    h: inches(1),
    text: 'Other',
  });
  expect(() => bringShapeToFront([shapes[0]!, other])).toThrow(/same parent/);
  expect(getSlideShapes(slide).map(getShapeText)).toEqual(['A', 'B', 'C', 'D', 'E']);
});
