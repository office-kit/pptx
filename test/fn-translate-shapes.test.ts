// translateShapes — bulk translate by (dx, dy).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeBounds,
  getSlides,
  inches,
  loadPresentation,
  translateShapes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: translateShapes', () => {
  it('shifts every shape by the same delta', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(3),
      y: inches(2),
      w: inches(1),
      h: inches(1),
    });

    const before = [getShapeBounds(a)!, getShapeBounds(b)!];
    translateShapes([a, b], 914400, -914400); // +1in, -1in

    const after = [getShapeBounds(a)!, getShapeBounds(b)!];
    expect(after[0]!.x).toBe(before[0]!.x + 914400);
    expect(after[0]!.y).toBe(before[0]!.y - 914400);
    expect(after[1]!.x).toBe(before[1]!.x + 914400);
    expect(after[1]!.y).toBe(before[1]!.y - 914400);
  });

  it('is a no-op on an empty list', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    void pres;
    expect(() => translateShapes([], 100, 100)).not.toThrow();
  });
});
