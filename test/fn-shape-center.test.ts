// getShapeCenter — center of the shape's bounds.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeCenter,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeCenter', () => {
  it('returns (x + w/2, y + h/2)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1), y: inches(1), w: inches(2), h: inches(4),
    });
    const center = getShapeCenter(rect)!;
    // x = 1in + 2in/2 = 2in = 1828800 EMU
    // y = 1in + 4in/2 = 3in = 2743200 EMU
    expect(center.x).toBe(914400 * 2);
    expect(center.y).toBe(914400 * 3);
  });
});
