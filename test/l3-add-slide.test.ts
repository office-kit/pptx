// Level-3 first feature: add a new slide from a chosen layout.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  duplicateSlide,
  importSlide,
  moveSlide,
  removeSlide,
  reverseSlides,
  isSlideHidden,
  setSlideHidden,
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
  it('keeps retained slide and shape edits live across deck changes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === 'Title Only')!;
    const first = addSlide(pres, { layout });
    const title = findSlidePlaceholder(first, 'title')!;
    const second = addSlide(pres, { layout });
    const assertRetainedEdits = async (label: string) => {
      setShapeText(title, label);
      setSlideHidden(first, true);
      expect(getSlides(pres)).toContain(first);
      expect(findSlidePlaceholder(first, 'title')).toBe(title);
      const reloaded = await loadPresentation(await savePresentation(pres));
      const index = getSlides(pres).indexOf(first);
      expect(isSlideHidden(getSlides(reloaded)[index]!)).toBe(true);
      expect(getShapeText(findSlidePlaceholder(getSlides(reloaded)[index]!, 'title')!)).toBe(label);
      setSlideHidden(first, false);
    };
    await assertRetainedEdits('After add / 追加後');
    const duplicate = duplicateSlide(pres, first);
    await assertRetainedEdits('After duplicate');
    moveSlide(pres, first, 2);
    await assertRetainedEdits('After move');
    reverseSlides(pres);
    await assertRetainedEdits('After reverse');
    removeSlide(pres, second);
    await assertRetainedEdits('After remove');
    expect(getSlides(pres)).not.toContain(second);
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    importSlide(pres, getSlides(source)[0]!, layout);
    await assertRetainedEdits('After import');
    expect(getShapeText(findSlidePlaceholder(duplicate, 'title')!)).toBe('After add / 追加後');
  });

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
    if (ctrTitle) setShapeText(ctrTitle, '@office-kit/pptx');
    const subTitle = findSlidePlaceholder(slide, 'subTitle');
    if (subTitle) setShapeText(subTitle, 'an OOXML library for TypeScript');

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reSlide = getSlides(reloaded)[0]!;
    const reCtr = findSlidePlaceholder(reSlide, 'ctrTitle');
    const reSub = findSlidePlaceholder(reSlide, 'subTitle');
    expect(reCtr && getShapeText(reCtr)).toBe('@office-kit/pptx');
    expect(reSub && getShapeText(reSub)).toBe('an OOXML library for TypeScript');
  });
});
