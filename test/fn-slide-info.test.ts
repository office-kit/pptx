// getSlideInfo — per-slide summary record.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getSlideInfo,
  getSlides,
  inches,
  loadPresentation,
  setSlideHidden,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideInfo', () => {
  it('captures index / title / hidden / shapeCount / layoutName', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    setSlideTitle(slide, 'Quarterly');
    setSlideHidden(slide, true);
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });

    const info = getSlideInfo(pres, slide);
    expect(info.index).toBe(0);
    expect(info.title).toBe('Quarterly');
    expect(info.hidden).toBe(true);
    expect(info.shapeCount).toBeGreaterThanOrEqual(1);
    expect(info.layoutName).toBe('Title and Content');
  });

  it('returns null fields gracefully when missing', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const info = getSlideInfo(pres, slide);
    expect(info.title).toBeNull();
    expect(info.hidden).toBe(false);
    expect(info.layoutName).toBe('Blank');
  });

  it('aggregates over an empty deck of slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const all = getSlides(pres).map((s) => getSlideInfo(pres, s));
    expect(all).toHaveLength(2);
    expect(all[0]!.index).toBe(0);
    expect(all[1]!.index).toBe(1);
  });
});
