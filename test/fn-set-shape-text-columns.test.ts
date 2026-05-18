// `setShapeTextColumns` — multi-column writer pairing the existing
// `getShapeTextColumns` reader.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeTextColumns,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextColumns,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeTextColumns', () => {
  it('round-trips count + gap through save/reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(2),
      text: 'col text',
    });
    setShapeTextColumns(tb, { count: 3, gapEmu: 228600 }); // 0.25in
    expect(getShapeTextColumns(tb)).toEqual({ count: 3, gapEmu: 228600 });

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeTextColumns(reShapes[reShapes.length - 1]!)).toEqual({
      count: 3,
      gapEmu: 228600,
    });
  });

  it('omits spcCol when gapEmu is not passed', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(2),
      text: 'c',
    });
    setShapeTextColumns(tb, { count: 2 });
    expect(getShapeTextColumns(tb)).toEqual({ count: 2 });
  });

  it('clears both attributes when set to null', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(2),
      text: 'c',
    });
    setShapeTextColumns(tb, { count: 4, gapEmu: 114300 });
    setShapeTextColumns(tb, null);
    expect(getShapeTextColumns(tb)).toBeNull();
  });

  it('throws when count < 2', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(6),
      h: inches(2),
      text: 'c',
    });
    expect(() => setShapeTextColumns(tb, { count: 1 })).toThrow(/count must be >= 2/);
  });
});
