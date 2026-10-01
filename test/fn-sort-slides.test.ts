// sortSlides — reorder slides by a custom comparator.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  getShapeText,
  getSlideShapes,
  moveSlide,
  getSlideText,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
  sortSlides,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const seedTitle = (
  pres: Awaited<ReturnType<typeof loadPresentation>>,
  slideIndex: number,
  value: string,
): void => {
  const slide = getSlides(pres)[slideIndex];
  if (!slide) return;
  const target = getSlideShapes(slide).find((s) => getShapeText(s).length > 0);
  if (!target) throw new Error(`no text shape on slide ${slideIndex}`);
  setShapeText(target, value);
};

describe('fn API: sortSlides', () => {
  it.each(['sort', 'move'] as const)(
    'retains absolute and normalized relationship targets during %s',
    async (operation) => {
      const parts = unzipSync(await readFile(fixture('two-slides.pptx')));
      const rels = 'ppt/_rels/presentation.xml.rels';
      parts[rels] = strToU8(
        strFromU8(parts[rels]!)
          .replace('Target="slides/slide1.xml"', 'Target="/ppt/slides/slide1.xml"')
          .replace('Target="slides/slide2.xml"', 'Target="./slides/slide2.xml"'),
      );
      const pres = await loadPresentation(zipSync(parts));
      const before = getSlides(pres);
      expect(before).toHaveLength(2);
      const rank = new Map(before.map((slide, index) => [slide, index]));
      if (operation === 'sort') sortSlides(pres, (a, b) => rank.get(b)! - rank.get(a)!);
      else moveSlide(pres, before[0]!, 1);
      expect(getSlides(pres)).toEqual([...before].reverse());
      const after = await loadPresentation(await savePresentation(pres));
      expect(getSlides(after).map(getSlideText)).toEqual([...before].reverse().map(getSlideText));
    },
  );

  it('reorders slides per the comparator', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    seedTitle(pres, 0, 'B-Second');
    seedTitle(pres, 1, 'A-First');

    sortSlides(pres, (a, b) => getSlideText(a).localeCompare(getSlideText(b)));

    const ordered = getSlides(pres).map((s) => getSlideText(s));
    expect(ordered[0]).toContain('A-First');
    expect(ordered[1]).toContain('B-Second');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reOrdered = getSlides(reloaded).map((s) => getSlideText(s));
    expect(reOrdered[0]).toContain('A-First');
    expect(reOrdered[1]).toContain('B-Second');
  });

  it('is a no-op when the comparator preserves the order', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getSlides(pres).map((s) => getSlideText(s));
    sortSlides(pres, () => 0);
    const after = getSlides(pres).map((s) => getSlideText(s));
    expect(after).toEqual(before);
  });
});
