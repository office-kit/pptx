// L3: duplicateSlide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  duplicateSlide,
  findSlidePlaceholder,
  getMediaParts,
  getShapeText,
  getSlideLayout,
  getSlideLayoutName,
  getSlideText,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: duplicateSlide', () => {
  it('produces a deep copy with independent text', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const original = getSlides(pres)[0]!;
    const originalText = getSlideText(original);
    const dup = duplicateSlide(pres, original);

    expect(getSlides(pres).length).toBe(3);
    expect(getSlideText(dup)).toBe(originalText);

    // Mutating the duplicate does not affect the original.
    const dupTitle = findSlidePlaceholder(dup, 'title');
    if (dupTitle) setShapeText(dupTitle, 'Duplicated slide');

    const slides = getSlides(pres);
    const origTitleAfter = findSlidePlaceholder(slides[0]!, 'title');
    expect(origTitleAfter && getShapeText(origTitleAfter)).toBe('Slide 1');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reSlides = getSlides(reloaded);
    expect(reSlides.length).toBe(3);
    expect(getShapeText(findSlidePlaceholder(reSlides[0]!, 'title')!)).toBe('Slide 1');
    expect(getShapeText(findSlidePlaceholder(reSlides[2]!, 'title')!)).toBe('Duplicated slide');
  });

  it('shares media parts with the source slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    duplicateSlide(pres, getSlides(pres)[0]!);
    const mediaParts = getMediaParts(pres);
    expect(mediaParts.length).toBe(1);
  });

  it('preserves layout binding on the duplicate', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const original = getSlides(pres)[0]!;
    const origLayoutName = (() => {
      const l = getSlideLayout(original);
      return l ? getSlideLayoutName(l) : null;
    })();
    const dup = duplicateSlide(pres, original);
    const dupLayout = getSlideLayout(dup);
    expect(dupLayout && getSlideLayoutName(dupLayout)).toBe(origLayoutName);
  });
});
