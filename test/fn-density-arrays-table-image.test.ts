// Density-array companions: `getPresentationTableCountsBySlide` and
// `getPresentationImageCountsBySlide`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  addSlideTable,
  getPresentationImageCountsBySlide,
  getPresentationTableCountsBySlide,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// Minimal 1×1 PNG (transparent).
const PNG_1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('density arrays: tables + images', () => {
  it('getPresentationTableCountsBySlide counts per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    addSlideTable(slideA!, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['A']],
    });
    addSlideTable(slideA!, {
      x: inches(0),
      y: inches(2),
      w: inches(4),
      h: inches(2),
      rows: [['B']],
    });
    addSlideTable(slideB!, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['C']],
    });
    expect(getPresentationTableCountsBySlide(pres)).toEqual([2, 1]);
  });

  it('getPresentationImageCountsBySlide counts per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    addSlideImage(slideA!, PNG_1x1, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    addSlideImage(slideB!, PNG_1x1, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    addSlideImage(slideB!, PNG_1x1, {
      x: inches(2),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    expect(getPresentationImageCountsBySlide(pres)).toEqual([1, 2]);
  });

  it('both return all-zeros on a clean deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationTableCountsBySlide(pres)).toEqual([0, 0]);
    expect(getPresentationImageCountsBySlide(pres)).toEqual([0, 0]);
  });
});
