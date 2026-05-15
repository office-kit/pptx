// findHyperlinkedShapes — shapes carrying external hyperlinks.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findHyperlinkedShapes,
  findSlideLayout,
  getShapeHyperlink,
  getShapeId,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findHyperlinkedShapes', () => {
  it('lists only shapes with an external hyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const linked = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'click',
    });
    setShapeHyperlink(linked, 'https://example.com');
    addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(2), h: inches(1), text: 'plain',
    });
    const out = findHyperlinkedShapes(slide);
    expect(out.length).toBe(1);
    expect(getShapeId(out[0]!)).toBe(getShapeId(linked));
    expect(getShapeHyperlink(out[0]!)).toBe('https://example.com');
  });

  it('returns empty when nothing is linked', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'a',
    });
    expect(findHyperlinkedShapes(slide)).toEqual([]);
  });
});
