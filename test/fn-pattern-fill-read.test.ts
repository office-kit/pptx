// `getShapePatternFill` — reads back the preset + fg/bg colors of a
// pattern fill, resolving theme tokens and color transforms via the
// shared color pipeline.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapePatternFill,
  getSlides,
  inches,
  loadPresentation,
  setShapePatternFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapePatternFill', () => {
  it('returns null when the shape has no pattern fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(3), h: inches(2),
    });
    expect(getShapePatternFill(pres, shape)).toBeNull();
  });

  it('round-trips preset + fg/bg colors', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(3), h: inches(2),
    });
    setShapePatternFill(shape, {
      preset: 'dkUpDiag',
      foreground: '#FF0000',
      background: '#0000FF',
    });
    const pat = getShapePatternFill(pres, shape);
    expect(pat).not.toBeNull();
    expect(pat!.preset).toBe('dkUpDiag');
    expect(pat!.foreground).toBe('#FF0000');
    expect(pat!.background).toBe('#0000FF');
  });

  it('resolves scheme color tokens via the theme', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(3), h: inches(2),
    });
    setShapePatternFill(shape, {
      preset: 'pct25',
      foreground: 'accent1',
      background: 'bg1',
    });
    const pat = getShapePatternFill(pres, shape);
    expect(pat).not.toBeNull();
    // Both should be hex resolved from the theme.
    expect(pat!.foreground).toMatch(/^#[0-9A-F]{6}$/);
    expect(pat!.background).toMatch(/^#[0-9A-F]{6}$/);
  });
});
