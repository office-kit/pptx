import { expect, it } from 'vitest';
import {
  copyShape,
  getGroupTransform,
  getShapeRotation,
  getShapeFlip,
  setShapeRotation,
  setShapeFlip,
  setShapeBounds,
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  groupShapes,
  getGroupChildren,
  getShapeText,
  getShapeZIndex,
  setShapeZIndex,
  bringShapeForward,
  sendShapeBackward,
  bringShapeToFront,
  sendShapeToBack,
  removeShape,
  getSlideShapes,
  getSlides,
  savePresentation,
  loadPresentation,
  inches,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { parseXml, NS } from '../src/internal/xml/index.ts';

it('reorders and removes nested shapes only among their siblings and preserves group extensions', async () => {
  const pres = createPresentation(),
    slide = addBlankSlide(pres);
  const [a, b, c, d] = ['A', 'B', 'C', 'Outside'].map((text) =>
    addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(1), h: inches(1), text }),
  );
  const group = groupShapes([a!, b!, c!]);
  const outer = groupShapes([group, d!]);
  group[SHAPE_ELEMENT].children.push(
    parseXml(`<p:extLst xmlns:p="${NS.pml}"><p:ext uri="keep"/></p:extLst>`).root,
  );
  const names = () => getGroupChildren(group).map(getShapeText);
  expect(getShapeZIndex(b!)).toBe(1);
  bringShapeForward(a!);
  expect(names()).toEqual(['B', 'A', 'C']);
  sendShapeBackward(a!);
  expect(names()).toEqual(['A', 'B', 'C']);
  bringShapeToFront(a!);
  expect(names()).toEqual(['B', 'C', 'A']);
  expect(group[SHAPE_ELEMENT].children.at(-1)).toMatchObject({ name: { localName: 'extLst' } });
  sendShapeToBack(a!);
  expect(names()).toEqual(['A', 'B', 'C']);
  setShapeZIndex(b!, 100);
  expect(names()).toEqual(['A', 'C', 'B']);
  expect(group[SHAPE_ELEMENT].children.at(-1)).toMatchObject({ name: { localName: 'extLst' } });
  removeShape(c!);
  expect(names()).toEqual(['A', 'B']);
  expect(getGroupChildren(outer)).toHaveLength(2);
  removeShape(c!);
  expect(names()).toEqual(['A', 'B']);
  setShapeZIndex(c!, 0);
  expect(names()).toEqual(['A', 'B']);
  const loaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  expect(getSlideShapes(loaded).map(getShapeText).filter(Boolean)).toEqual(['A', 'B', 'Outside']);
});

it('copies a nested object with ancestor transforms without copying its siblings', async () => {
  const pres = createPresentation(),
    slide = addBlankSlide(pres);
  const [a, b, c] = ['日本語', 'Sibling', 'Outside'].map((text) =>
    addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(2), h: inches(1), text }),
  );
  const inner = groupShapes([a!, b!]);
  setShapeRotation(inner, 37);
  setShapeFlip(inner, { horizontal: true, vertical: false });
  const outer = groupShapes([inner, c!]);
  setShapeRotation(outer, 23);
  setShapeBounds(outer, { x: inches(3), y: inches(2), w: inches(6), h: inches(2) });
  const destination = addBlankSlide(pres);
  const copy = copyShape(destination, a!, { preserveGroupTransform: true });
  const copiedInner = getGroupChildren(copy)[0]!;
  expect(getGroupTransform(copy)).toEqual(getGroupTransform(outer));
  expect(getShapeRotation(copy)).toBe(23);
  expect(getGroupTransform(copiedInner)).toEqual(getGroupTransform(inner));
  expect(getShapeRotation(copiedInner)).toBe(37);
  expect(getShapeFlip(copiedInner)).toEqual(getShapeFlip(inner));
  expect(getSlideShapes(destination).map(getShapeText).filter(Boolean)).toEqual(['日本語']);
  expect(getGroupChildren(inner)).toHaveLength(2);
  expect(getGroupChildren(outer)).toHaveLength(2);
  const loaded = await loadPresentation(await savePresentation(pres));
  expect(getSlideShapes(getSlides(loaded)[1]!).map(getShapeText).filter(Boolean)).toEqual([
    '日本語',
  ]);
});
