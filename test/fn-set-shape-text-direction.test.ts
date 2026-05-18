// `setShapeTextDirection` — vertical-text writer that pairs with
// `getShapeTextDirection`. Asserts all six `ST_TextVerticalType`
// values round-trip and that null / `'horz'` clears the attribute.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeTextDirection,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextDirection,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const DIRECTIONS = [
  'vert',
  'vert270',
  'wordArtVert',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
] as const;

describe('fn API: setShapeTextDirection', () => {
  it('round-trips every non-default direction in-memory', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'v',
    });
    for (const dir of DIRECTIONS) {
      setShapeTextDirection(tb, dir);
      expect(getShapeTextDirection(tb)).toBe(dir);
    }
  });

  it('round-trips through save/reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'v',
    });
    setShapeTextDirection(tb, 'eaVert');

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeTextDirection(reShapes[reShapes.length - 1]!)).toBe('eaVert');
  });

  it('clears the attribute when set to null or horz', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'v',
    });
    setShapeTextDirection(tb, 'vert');
    expect(getShapeTextDirection(tb)).toBe('vert');
    setShapeTextDirection(tb, null);
    expect(getShapeTextDirection(tb)).toBeNull();
    setShapeTextDirection(tb, 'vert');
    setShapeTextDirection(tb, 'horz');
    expect(getShapeTextDirection(tb)).toBeNull();
  });
});
