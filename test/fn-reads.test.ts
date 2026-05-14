// Free-function read API: slide + shape accessors.
//
// Verifies that the tree-shakeable `fn.ts` query functions return the
// same values as the equivalent class-API getters, on real
// `python-pptx` fixtures. Pairs the fn-API readout against the class
// API as an oracle so behavior drift is loud.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  findSlidePlaceholder,
  getShapeFlip,
  getShapeId,
  getShapeKind,
  getShapeName,
  getShapePlaceholderIdx,
  getShapePlaceholderType,
  getShapePosition,
  getShapeRotation,
  getShapeSize,
  getShapeText,
  getSlideLayout,
  getSlideShapes,
  getSlideText,
  getSlides,
  loadPresentation,
  replaceTokensInSlide,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide reads', () => {
  it('getSlideShapes / getSlideLayout / findSlidePlaceholder agree with class API', async () => {
    const bytes = await readFile(fixture('one-text-slide.pptx'));
    const fnPres = await loadPresentation(bytes);
    const clsPres = await Presentation.load(bytes);

    const fnSlide = getSlides(fnPres)[0]!;
    const clsSlide = clsPres.slides[0]!;

    expect(getSlideShapes(fnSlide).length).toBe(clsSlide.shapes.length);
    expect(getSlideText(fnSlide)).toBe(clsSlide.text);
    // SlideLayoutData uses internal Symbol keys — not addressable from
    // tests. Confirm presence/absence parity with the class API.
    expect(getSlideLayout(fnSlide) === null).toBe(clsSlide.layout === null);

    // findSlidePlaceholder for the "title" placeholder must match.
    const fnTitle = findSlidePlaceholder(fnSlide, 'title');
    const clsTitle = clsSlide.findPlaceholder('title');
    expect((fnTitle === null) === (clsTitle === null)).toBe(true);
    if (fnTitle && clsTitle) {
      expect(getShapeText(fnTitle)).toBe(clsTitle.text);
      expect(getShapeName(fnTitle)).toBe(clsTitle.name);
      expect(getShapeId(fnTitle)).toBe(clsTitle.id);
      expect(getShapeKind(fnTitle)).toBe(clsTitle.kind);
      expect(getShapePlaceholderType(fnTitle)).toBe(clsTitle.placeholderType);
      expect(getShapePlaceholderIdx(fnTitle)).toBe(clsTitle.placeholderIdx);
    }
  });

  it('getShape{Position,Size,Rotation,Flip} return the same EMU values as class API', async () => {
    const bytes = await readFile(fixture('one-image-slide.pptx'));
    const fnPres = await loadPresentation(bytes);
    const clsPres = await Presentation.load(bytes);

    const fnSlide = getSlides(fnPres)[0]!;
    const clsSlide = clsPres.slides[0]!;
    const fnShapes = getSlideShapes(fnSlide);
    expect(fnShapes.length).toBe(clsSlide.shapes.length);

    for (let i = 0; i < fnShapes.length; i++) {
      const fnShape = fnShapes[i]!;
      const clsShape = clsSlide.shapes[i]!;
      expect(getShapePosition(fnShape)).toEqual(clsShape.position);
      expect(getShapeSize(fnShape)).toEqual(clsShape.size);
      expect(getShapeRotation(fnShape)).toBe(clsShape.rotation);
      expect(getShapeFlip(fnShape)).toEqual(clsShape.flip);
    }
  });

  it('replaceTokensInSlide mutates a single slide in place', async () => {
    const bytes = await readFile(fixture('one-text-slide.pptx'));
    // Seed a token by mutating via the class API.
    const seedPres = await Presentation.load(bytes);
    const seedShape = seedPres.slides[0]?.shapes.find((s) => s.text.length > 0);
    if (!seedShape) throw new Error('expected text shape');
    seedShape.setText('Hello, {{name}}!');
    const seededBytes = await seedPres.save();

    // Reload via fn API and replace via fn API.
    const pres = await loadPresentation(seededBytes);
    const slide = getSlides(pres)[0]!;
    const n = replaceTokensInSlide(slide, { name: 'World' });
    expect(n).toBe(1);
    expect(getSlideText(slide)).toContain('Hello, World!');

    // Persists across round-trip.
    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides[0]?.text).toContain('Hello, World!');
  });
});
