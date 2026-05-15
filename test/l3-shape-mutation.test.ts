// L3: shape-level mutations (setPosition / setSize / remove) plus an
// end-to-end "build a deck from scratch" smoke test that exercises
// every authoring primitive we ship.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideImage,
  addSlideTextBox,
  findSlideLayout,
  findSlidePlaceholder,
  getShapeKind,
  getShapePlaceholderIdx,
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
  removeShape,
  removeSlide,
  replaceTokensInPresentation,
  savePresentation,
  setShapePosition,
  setShapeSize,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('L3: setShapePosition / setShapeSize', () => {
  it('moves and resizes an existing picture shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture');
    if (!picture) throw new Error('expected picture');

    setShapePosition(picture, inches(4), inches(2));
    setShapeSize(picture, inches(1), inches(1));

    const reloaded = await loadPresentation(await savePresentation(pres));
    const repic = getSlideShapes(getSlides(reloaded)[0]!).find((s) => getShapeKind(s) === 'picture');
    expect(repic && getShapePosition(repic)).toEqual({ x: inches(4), y: inches(2) });
    expect(repic && getShapeSize(repic)).toEqual({ w: inches(1), h: inches(1) });
  });

  it('creates a transform on a layout-inheriting placeholder when set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const title = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!title) throw new Error('expected title');
    expect(getShapePosition(title)).toBeNull();
    setShapePosition(title, inches(1), inches(5));
    setShapeSize(title, inches(6), inches(0.5));

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reTitle = findSlidePlaceholder(getSlides(reloaded)[0]!, 'title');
    expect(reTitle && getShapePosition(reTitle)).toEqual({ x: inches(1), y: inches(5) });
    expect(reTitle && getShapeSize(reTitle)).toEqual({ w: inches(6), h: pt(36) });
  });
});

describe('L3: removeShape', () => {
  it('removes a shape and shrinks the shape tree by one', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const initialCount = getSlideShapes(slide).length;
    const body = getSlideShapes(slide).find((s) => getShapePlaceholderIdx(s) === 1);
    if (!body) throw new Error('expected body placeholder');
    removeShape(body);

    expect(getSlideShapes(getSlides(pres)[0]!).length).toBe(initialCount - 1);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(reloaded)[0]!).length).toBe(initialCount - 1);
  });
});

describe('L3: end-to-end deck build from blank', () => {
  it('combines every authoring primitive into a single working deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));

    const titleLayout = findSlideLayout(pres, 'Title Slide');
    if (!titleLayout) throw new Error('expected Title Slide layout');
    const slide1 = addSlide(pres, { layout: titleLayout });
    {
      const ctr = findSlidePlaceholder(slide1, 'ctrTitle');
      if (ctr) setShapeText(ctr, 'pptx-kit demo');
      const sub = findSlidePlaceholder(slide1, 'subTitle');
      if (sub) setShapeText(sub, 'end-to-end deck from a blank template');
    }

    const blank = findSlideLayout(pres, 'Blank');
    if (!blank) throw new Error('expected Blank layout');
    const slide2 = addSlide(pres, { layout: blank });
    const heading = addSlideTextBox(slide2, {
      x: inches(1), y: inches(1), w: inches(8), h: inches(0.7), text: 'Body section heading',
    });
    setShapeSize(heading, inches(8), inches(0.8));
    const body = addSlideTextBox(slide2, {
      x: inches(1), y: inches(2), w: inches(8), h: inches(3), text: 'Line 1\nLine 2\nLine 3',
    });
    setShapeText(body, 'Replaced body text');
    addSlideImage(slide2, PNG_1X1, { x: inches(1), y: inches(5), w: inches(2), h: inches(2) });

    const sectionLayout = findSlideLayout(pres, 'Section Header');
    if (!sectionLayout) throw new Error('expected Section Header layout');
    const slide3 = addSlide(pres, { layout: sectionLayout });
    removeSlide(pres, slide3);

    // Re-fetch handles before the move (cache was invalidated).
    const slides = getSlides(pres);
    moveSlide(pres, slides[1]!, 0);

    // Replace a token across the deck.
    {
      const slidesAfter = getSlides(pres);
      const ctr = findSlidePlaceholder(slidesAfter[1]!, 'ctrTitle');
      if (ctr) setShapeText(ctr, 'Hello, {{name}}');
    }
    const n = replaceTokensInPresentation(pres, { name: 'Alice' });
    expect(n).toBe(1);

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reSlides = getSlides(reloaded);

    expect(reSlides.length).toBe(2);

    const slideA = reSlides[0]!;
    expect(getSlideText(slideA)).toContain('Body section heading');
    expect(getSlideText(slideA)).toContain('Replaced body text');
    expect(getSlideShapes(slideA).some((s) => getShapeKind(s) === 'picture')).toBe(true);

    const slideB = reSlides[1]!;
    const reCtr = findSlidePlaceholder(slideB, 'ctrTitle');
    expect(reCtr && getShapeText(reCtr)).toBe('Hello, Alice');
  });
});
