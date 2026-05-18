// `findShapesWithHyperlinks(slide)` — every shape with any hyperlink.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapesWithHyperlinks,
  getSlides,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesWithHyperlinks', () => {
  it('returns every shape with any hyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    const b = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'B',
    });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'C',
    });
    setShapeHyperlink(a, 'https://a.example/');
    setShapeHyperlink(b, 'https://b.example/');
    const matches = findShapesWithHyperlinks(slide);
    expect(matches.length).toBe(2);
  });

  it('returns an empty array on a slide with no hyperlinks', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(findShapesWithHyperlinks(slide)).toEqual([]);
  });
});
