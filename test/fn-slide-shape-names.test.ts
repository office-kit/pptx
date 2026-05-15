// getSlideShapeNames — list every shape's cNvPr@name.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getSlideShapeNames,
  getSlideShapes,
  getShapeName,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideShapeNames', () => {
  it('matches the names from getSlideShapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1),
      text: 'a',
      name: 'CustomTagOne',
    });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(1), h: inches(1),
      text: 'b',
      name: 'CustomTagTwo',
    });
    const fromBuild = getSlideShapes(slide).map((s) => getShapeName(s));
    expect(getSlideShapeNames(slide)).toEqual(fromBuild);
    const names = getSlideShapeNames(slide);
    expect(names).toContain('CustomTagOne');
    expect(names).toContain('CustomTagTwo');
  });
});
