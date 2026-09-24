// getShapeZIndex + setShapeZIndex.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  createPresentation,
  addBlankSlide,
  groupShapes,
  getGroupChildren,
  getShapeId,
  bringShapeToFront,
  bringShapeForward,
  sendShapeToBack,
  sendShapeBackward,
  removeShape,
  savePresentation,
  getShapeName,
  getShapeZIndex,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  setShapeZIndex,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const orderedNames = (slide: ReturnType<typeof getSlides>[number]): string[] =>
  getSlideShapes(slide).map((s) => getShapeName(s));

describe('fn API: shape z-index', () => {
  it('getShapeZIndex returns the document-order position', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'B',
    });
    expect(getShapeZIndex(a)).toBeGreaterThan(-1);
    expect(getShapeZIndex(b)).toBeGreaterThan(getShapeZIndex(a));
  });

  it('setShapeZIndex moves the shape to the requested position', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'B',
    });
    const c = addSlideShape(slide, {
      preset: 'triangle',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'C',
    });

    setShapeZIndex(c, 0);
    expect(orderedNames(slide)[0]).toBe('C');
  });

  it('setShapeZIndex clamps to the available range', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    setShapeZIndex(a, 999);
    // After clamp the shape sits at the last position among shapes.
    expect(orderedNames(slide).at(-1)).toBe('A');
  });
});

it('orders and deletes descendants within their own group', async () => {
  const pres = createPresentation(),
    slide = addBlankSlide(pres);
  const shapes = ['Outside', 'A', 'B', 'C'].map((name) =>
    addSlideShape(slide, {
      name,
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    }),
  );
  const [outside, a, b, c] = shapes;
  const group = groupShapes([c!, a!, b!]);
  const names = () => getGroupChildren(group).map(getShapeName);
  expect(names()).toEqual(['A', 'B', 'C']);
  expect(getShapeZIndex(b!)).toBe(1);
  bringShapeToFront(a!);
  expect(names()).toEqual(['B', 'C', 'A']);
  sendShapeToBack(a!);
  expect(names()).toEqual(['A', 'B', 'C']);
  bringShapeForward(a!);
  expect(names()).toEqual(['B', 'A', 'C']);
  sendShapeBackward(a!);
  expect(names()).toEqual(['A', 'B', 'C']);
  setShapeZIndex(c!, 0);
  expect(names()).toEqual(['C', 'A', 'B']);
  removeShape(a!);
  expect(names()).toEqual(['C', 'B']);
  expect(getShapeZIndex(a!)).toBe(-1);
  setShapeZIndex(a!, 0); // stale removed handle must not reattach itself
  expect(names()).toEqual(['C', 'B']);
  expect(getShapeZIndex(outside!)).toBe(0);
  expect(getShapeZIndex(group)).toBe(1);
  const loaded = await loadPresentation(await savePresentation(pres));
  const loadedShapes = getSlideShapes(getSlides(loaded)[0]!);
  const savedGroup = loadedShapes.find((s) => getShapeId(s) === getShapeId(group))!;
  expect(getGroupChildren(savedGroup).map(getShapeName)).toEqual(['C', 'B']);
  expect(loadedShapes.some((s) => getShapeId(s) === getShapeId(a!))).toBe(false);
});
