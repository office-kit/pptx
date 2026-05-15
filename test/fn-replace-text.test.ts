// Free-text replace across the deck (no {{token}} braces required).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeText,
  getSlideShapes,
  getSlideText,
  getSlides,
  loadPresentation,
  replaceTextInPresentation,
  replaceTextInSlide,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const seedTextOnFirstTextShape = async (
  fixtureName: string,
  values: ReadonlyArray<string>,
): Promise<Uint8Array> => {
  const pres = await loadPresentation(await readFile(fixture(fixtureName)));
  const slides = getSlides(pres);
  for (let i = 0; i < values.length; i++) {
    const slide = slides[i];
    if (!slide) break;
    const target = getSlideShapes(slide).find((s) => getShapeText(s).length > 0);
    if (!target) throw new Error(`expected a text shape on slide ${i}`);
    setShapeText(target, values[i]!);
  }
  return savePresentation(pres);
};

describe('fn API: free-text replace', () => {
  it('replaceTextInPresentation rewrites a literal across every slide', async () => {
    const bytes = await seedTextOnFirstTextShape('one-text-slide.pptx', [
      'Hello FOO and FOO again',
    ]);
    const seeded = await loadPresentation(bytes);
    const n = replaceTextInPresentation(seeded, 'FOO', 'BAR');
    expect(n).toBeGreaterThan(0);
    expect(getSlideText(getSlides(seeded)[0]!)).toContain('Hello BAR and BAR again');
  });

  it('replaceTextInPresentation accepts a RegExp', async () => {
    const bytes = await seedTextOnFirstTextShape('one-text-slide.pptx', [
      'order #12345 and #67890',
    ]);
    const seeded = await loadPresentation(bytes);
    const n = replaceTextInPresentation(seeded, /#\d+/, 'REDACTED');
    expect(n).toBeGreaterThan(0);
    expect(getSlideText(getSlides(seeded)[0]!)).toContain('REDACTED');
    expect(getSlideText(getSlides(seeded)[0]!)).not.toContain('#12345');
  });

  it('replaceTextInSlide is slide-scoped', async () => {
    const bytes = await seedTextOnFirstTextShape('two-slides.pptx', [
      'slide-one says A',
      'slide-two says A',
    ]);
    const seeded = await loadPresentation(bytes);

    const slides = getSlides(seeded);
    const n = replaceTextInSlide(slides[0]!, 'A', 'BB');
    expect(n).toBeGreaterThan(0);
    expect(getSlideText(slides[0]!)).toContain('BB');
    expect(getSlideText(slides[1]!)).toContain('says A'); // not touched

    expect(getSlideShapes(slides[0]!).length).toBeGreaterThan(0);
  });
});
