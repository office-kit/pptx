// findSlideByText / findSlidesByText.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlideByText,
  findSlidesByText,
  getShapeText,
  getSlideIndex,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const seedFirstShapeOn = (
  pres: Parameters<typeof getSlides>[0],
  slideIndex: number,
  value: string,
): void => {
  const slide = getSlides(pres)[slideIndex];
  if (!slide) return;
  const target = getSlideShapes(slide).find((s) => getShapeText(s).length > 0);
  if (!target) throw new Error(`no text shape on slide ${slideIndex}`);
  setShapeText(target, value);
};

describe('fn API: findSlideByText / findSlidesByText', () => {
  it('finds the first slide containing the needle', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    seedFirstShapeOn(pres, 0, 'Brand: Acme');
    seedFirstShapeOn(pres, 1, 'Brand: Acme again');

    const first = findSlideByText(pres, 'Acme');
    expect(first).not.toBeNull();
    expect(getSlideIndex(pres, first!)).toBe(0);

    const all = findSlidesByText(pres, 'Acme');
    expect(all).toHaveLength(2);
  });

  it('accepts a RegExp', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    seedFirstShapeOn(pres, 0, 'Order #12345');
    const found = findSlideByText(pres, /#\d{5}/);
    expect(found).not.toBeNull();
  });

  it('returns null / empty when no slide matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(findSlideByText(pres, 'NO_SUCH_TEXT_xyz123')).toBeNull();
    expect(findSlidesByText(pres, 'NO_SUCH_TEXT_xyz123')).toEqual([]);
  });
});
