// Level-3 first feature: add a new slide from a chosen layout.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlidePlaceholder,
  getShapeText,
  getSlideLayout,
  getSlideLayoutName,
  getSlideLayouts,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: addSlide from a layout', () => {
  it('adds a single slide bound to the chosen layout and persists', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getSlides(pres).length).toBe(0);

    const layout = getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === 'Title and Content');
    if (!layout) throw new Error('expected Title and Content layout');

    const slide = addSlide(pres, { layout });
    expect(slide).toBeDefined();
    expect(getSlides(pres).length).toBe(1);

    const title = findSlidePlaceholder(slide, 'title');
    expect(title).not.toBeNull();
    if (title) setShapeText(title, 'Brand new slide');

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(reloaded).length).toBe(1);
    const reSlide = getSlides(reloaded)[0]!;
    const reTitle = findSlidePlaceholder(reSlide, 'title');
    expect(reTitle).not.toBeNull();
    expect(reTitle && getShapeText(reTitle)).toBe('Brand new slide');
    const reLayout = getSlideLayout(reSlide);
    expect(reLayout && getSlideLayoutName(reLayout)).toBe('Title and Content');
  });

  it('preserves existing slides and appends in order', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const baseline = getSlides(pres).length;
    const layout = getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === 'Title Only');
    if (!layout) throw new Error('expected Title Only layout');
    addSlide(pres, { layout });
    addSlide(pres, { layout });

    expect(getSlides(pres).length).toBe(baseline + 2);
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(reloaded).length).toBe(baseline + 2);
  });

  it('builds blank.pptx into a usable single-slide deck end-to-end', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === 'Title Slide');
    if (!layout) throw new Error('expected Title Slide layout');
    const slide = addSlide(pres, { layout });
    const ctrTitle = findSlidePlaceholder(slide, 'ctrTitle');
    if (ctrTitle) setShapeText(ctrTitle, 'pptx-kit');
    const subTitle = findSlidePlaceholder(slide, 'subTitle');
    if (subTitle) setShapeText(subTitle, 'an OOXML library for TypeScript');

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reSlide = getSlides(reloaded)[0]!;
    const reCtr = findSlidePlaceholder(reSlide, 'ctrTitle');
    const reSub = findSlidePlaceholder(reSlide, 'subTitle');
    expect(reCtr && getShapeText(reCtr)).toBe('pptx-kit');
    expect(reSub && getShapeText(reSub)).toBe('an OOXML library for TypeScript');
  });
});
