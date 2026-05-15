// Level-3: remove a slide from the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlidePlaceholder,
  getShapeText,
  getSlideLayoutName,
  getSlideLayouts,
  getSlides,
  loadPresentation,
  removeSlide,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: removeSlide', () => {
  it('removes the first slide and renumbers nothing', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlides(pres).length).toBe(2);
    const secondSlide = getSlides(pres)[1]!;
    const secondTitle = (() => {
      const ph = findSlidePlaceholder(secondSlide, 'title');
      return ph ? getShapeText(ph) : undefined;
    })();

    const first = getSlides(pres)[0]!;
    removeSlide(pres, first);
    expect(getSlides(pres).length).toBe(1);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(reloaded).length).toBe(1);
    const reTitle = findSlidePlaceholder(getSlides(reloaded)[0]!, 'title');
    expect(reTitle && getShapeText(reTitle)).toBe(secondTitle);
  });

  it('round-trips an empty deck after removing the only slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    expect(getSlides(pres).length).toBe(1);
    const only = getSlides(pres)[0]!;
    removeSlide(pres, only);
    expect(getSlides(pres).length).toBe(0);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(reloaded).length).toBe(0);
  });

  it('does NOT reuse the freed sldId on the next addSlide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const first = getSlides(pres)[0]!;
    removeSlide(pres, first);

    const layout = getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === 'Title Only');
    if (!layout) throw new Error('expected layout');
    addSlide(pres, { layout });

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(reloaded).length).toBe(2);
  });

  it('handles a slide from a different presentation by matching part name', async () => {
    const pres1 = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const pres2 = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres1)[0]!;
    removeSlide(pres1, slide);
    expect(getSlides(pres1).length).toBe(1);
    // pres2 is independent
    expect(getSlides(pres2).length).toBe(2);
  });
});
