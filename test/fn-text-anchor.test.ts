// Vertical text anchor on a shape's bodyPr.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextAnchor,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  const slide = getSlides(pres)[slideIndex];
  if (!slide) throw new Error(`slide ${slideIndex} not found`);
  return getSlideXmlString(slide);
};

describe('fn API: setShapeTextAnchor', () => {
  it('writes anchor="t" / "ctr" / "b" to <a:bodyPr>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(2), text: 'A',
    });
    const b = addSlideTextBox(slide, {
      x: inches(2), y: inches(0), w: inches(2), h: inches(2), text: 'B',
    });
    const c = addSlideTextBox(slide, {
      x: inches(4), y: inches(0), w: inches(2), h: inches(2), text: 'C',
    });
    setShapeTextAnchor(a, 'top');
    setShapeTextAnchor(b, 'center');
    setShapeTextAnchor(c, 'bottom');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('anchor="t"');
    expect(xml).toContain('anchor="ctr"');
    expect(xml).toContain('anchor="b"');
  });

  it('replaces an existing anchor on subsequent calls', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(2), text: 'A',
    });
    setShapeTextAnchor(tb, 'top');
    setShapeTextAnchor(tb, 'bottom');
    const xml = await slideXml(await savePresentation(pres), 0);
    const matches = xml.match(/<a:bodyPr[^>]*anchor="[^"]+"/g) ?? [];
    expect(matches.length).toBeGreaterThan(0);
    expect(matches.every((m) => m.endsWith('"b"'))).toBe(true);
  });
});
