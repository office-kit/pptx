// sortSlides — reorder slides by a custom comparator.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeText,
  getSlideShapes,
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
