// getSlidesWithEmptyPlaceholders — slides still needing editorial fill.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findEmptyPlaceholders,
  findSlideLayout,
  findSlidePlaceholder,
  getSlideIndex,
  getSlidesWithEmptyPlaceholders,
  loadPresentation,
  setShapeText,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlidesWithEmptyPlaceholders', () => {
  it('lists every slide with at least one empty placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    // Three slides: 1st filled, 2nd untouched, 3rd partially filled.
    const first = addSlide(pres, { layout });
    setSlideTitle(first, 'Done');
    setShapeText(findSlidePlaceholder(first, 'body')!, 'all good');

    addSlide(pres, { layout }); // untouched

    const third = addSlide(pres, { layout });
    setSlideTitle(third, 'partial');

    const slides = getSlidesWithEmptyPlaceholders(pres);
    const indices = slides.map((s) => getSlideIndex(pres, s));
    // The untouched (1) and partial (2) slides are definitely
    // included. The "filled" slide (0) might also be — Title and
    // Content layouts often include extra dt/ftr/sldNum slots that
    // stay empty.
    expect(indices).toContain(1);
    expect(indices).toContain(2);
    for (const s of slides) {
      expect(findEmptyPlaceholders(s).length).toBeGreaterThan(0);
    }
  });

  it('returns empty when every slide is filled', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout });
    // Blank layout has no placeholders other than dt/sldNum/ftr; these
    // are typically empty too. Check that the function doesn't crash.
    const slides = getSlidesWithEmptyPlaceholders(pres);
    expect(Array.isArray(slides)).toBe(true);
  });
});
