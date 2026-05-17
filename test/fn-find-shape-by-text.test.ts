// findShapeByText / findShapesByText — locate shapes by their
// visible text. Complement to findShapeByName.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapeByText,
  findShapesByText,
  getShapeText,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapeByText / findShapesByText', () => {
  it('finds the first shape containing the substring', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'Quarterly Results',
    });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Annual Report',
    });

    const hit = findShapeByText(slide, 'Quarterly');
    expect(hit).not.toBeNull();
    expect(getShapeText(hit!)).toContain('Quarterly');
  });

  it('accepts a RegExp', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'Revenue $1.2M',
    });
    const hit = findShapeByText(slide, /\$\d+(?:\.\d+)?M/);
    expect(hit).not.toBeNull();
  });

  it('returns null when no shape matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(findShapeByText(slide, 'no-such-text-anywhere')).toBeNull();
  });

  it('findShapesByText returns every match', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'TODO: write',
    });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'TODO: edit',
    });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'Done',
    });
    const matches = findShapesByText(slide, 'TODO');
    expect(matches.length).toBe(2);
  });
});
