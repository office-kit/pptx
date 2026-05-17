// getShapeXmlString — diagnostic dump of a single shape's XML.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeXmlString,
  getSlides,
  inches,
  loadPresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeXmlString', () => {
  it("returns the shape's XML body", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'unique-shape-xml-marker',
    });
    const xml = getShapeXmlString(tb);
    expect(xml).toContain('unique-shape-xml-marker');
    expect(xml.startsWith('<')).toBe(true);
  });

  it('reflects pending edits', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'before',
    });
    const before = getShapeXmlString(tb);
    setShapeText(tb, 'after');
    const after = getShapeXmlString(tb);
    expect(after).toContain('after');
    expect(after).not.toBe(before);
  });
});
