// Level-3: moveSlide + addSlideTextBox.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  findSlidePlaceholder,
  getShapeKind,
  getShapeName,
  getShapePosition,
  getShapeSize,
  getShapeText,
  getSlideShapes,
  getSlideText,
  getSlides,
  inches,
  loadPresentation,
  moveSlide,
  pt,
  removeSlide,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const titleOf = (slide: ReturnType<typeof getSlides>[number]): string | undefined => {
  const ph = findSlidePlaceholder(slide, 'title');
  return ph ? getShapeText(ph) : undefined;
};

describe('L3: moveSlide', () => {
  it('reorders two slides and persists the new order', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    expect(titleOf(slides[0]!)).toBe('Slide 1');
    expect(titleOf(slides[1]!)).toBe('Slide 2');

    moveSlide(pres, slides[0]!, 1);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reSlides = getSlides(reloaded);
    expect(titleOf(reSlides[0]!)).toBe('Slide 2');
    expect(titleOf(reSlides[1]!)).toBe('Slide 1');
  });

  it('clamps out-of-range indices', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    moveSlide(pres, getSlides(pres)[0]!, 99);
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(titleOf(getSlides(reloaded)[1]!)).toBe('Slide 1');
  });

  it('move + remove + add interleave round-trips cleanly', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(pres, 'Title Only');
    if (!layout) throw new Error('expected layout');
    addSlide(pres, { layout });
    moveSlide(pres, getSlides(pres)[2]!, 0);
    removeSlide(pres, getSlides(pres)[2]!);
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(reloaded).length).toBe(2);
  });
});

describe('L3: addSlideTextBox', () => {
  it('appends a free-form text box to a slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const initialShapeCount = getSlideShapes(slide).length;

    const box = addSlideTextBox(slide, {
      x: inches(1), y: inches(1), w: inches(4), h: inches(1), text: 'Hello from addTextBox',
    });

    expect(getShapeText(box)).toBe('Hello from addTextBox');
    expect(getShapeKind(box)).toBe('shape');
    expect(getSlideShapes(getSlides(pres).at(-1)!).length).toBe(initialShapeCount + 1);

    expect(getShapePosition(box)).toEqual({ x: inches(1), y: inches(1) });
    expect(getShapeSize(box)).toEqual({ w: inches(4), h: inches(1) });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reSlide = getSlides(reloaded)[0]!;
    expect(getSlideText(reSlide)).toContain('Hello from addTextBox');
  });

  it('supports a custom shape name', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(2), y: pt(50), w: inches(3), h: pt(40),
      text: 'named', name: 'My Custom Box',
    });
    expect(getShapeName(box)).toBe('My Custom Box');
  });
});
