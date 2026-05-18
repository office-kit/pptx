// `findShapesByName(slide, RegExp)` — pattern-match variant.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapesByName,
  getShapeName,
  getSlides,
  inches,
  loadPresentation,
  renameShape,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesByName (RegExp)', () => {
  it('matches every shape whose name fits the pattern', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    for (const i of [1, 2, 3]) {
      const tb = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(i - 1),
        w: inches(2),
        h: inches(1),
        text: `t${i}`,
      });
      renameShape(tb, `BrandPlaceholder${i}`);
    }
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(4),
      w: inches(2),
      h: inches(1),
      text: 'ignored',
    });
    const matches = findShapesByName(slide, /^BrandPlaceholder\d+$/);
    expect(matches.length).toBe(3);
    expect(matches.every((m) => /^BrandPlaceholder\d+$/.test(getShapeName(m)))).toBe(true);
  });

  it('still matches exact strings (backwards-compat)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 't',
    });
    renameShape(tb, 'OnlyMe');
    const matches = findShapesByName(slide, 'OnlyMe');
    expect(matches.length).toBe(1);
  });
});
