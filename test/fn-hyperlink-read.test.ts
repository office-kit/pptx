// getShapeHyperlink — read back the text hyperlink set via setShapeHyperlink.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeHyperlink,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeHyperlink', () => {
  it('returns null when no run carries a hyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide).find((s) => s !== undefined)!;
    expect(getShapeHyperlink(shape)).toBeNull();
  });

  it('round-trips a URL', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide).find((s) => s !== undefined)!;
    setShapeHyperlink(shape, 'https://example.com/');
    expect(getShapeHyperlink(shape)).toBe('https://example.com/');
    setShapeHyperlink(shape, null);
    expect(getShapeHyperlink(shape)).toBeNull();
  });
});
