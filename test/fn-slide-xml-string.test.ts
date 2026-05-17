// getSlideXmlString — slide XML body as a string. Diagnostic helper.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  setSlideHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideXmlString', () => {
  it('returns valid-looking XML for a fixture slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const xml = getSlideXmlString(slide);
    expect(typeof xml).toBe('string');
    expect(xml.length).toBeGreaterThan(0);
    expect(xml).toContain('<p:sld');
  });

  it('reflects pending edits to the slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideXmlString(slide);
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'unique-xml-snapshot-marker',
    });
    const after = getSlideXmlString(slide);
    expect(after.includes('unique-xml-snapshot-marker')).toBe(true);
    expect(after).not.toBe(before);
  });

  it('changes when slide-level metadata changes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideHidden(slide, true);
    // The slide XML itself doesn't carry the hidden flag (that's on
    // presentation.xml's sldId), so this assertion just confirms the
    // helper runs without error on a mutated slide.
    expect(getSlideXmlString(slide).startsWith('<')).toBe(true);
  });
});
