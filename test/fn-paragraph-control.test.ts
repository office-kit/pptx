// Per-paragraph alignment + nesting level.
//
// Pairs with the run-level helpers to give callers full control over a
// shape's text body without rewriting it from scratch.

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
  setParagraphAlignment,
  setParagraphLevel,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: per-paragraph control', () => {
  it('setParagraphAlignment targets one paragraph only', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'first\nsecond\nthird',
    });
    setParagraphAlignment(tb, 1, 'center');
    const xml = await slideXml(await savePresentation(pres), 0);
    // Exactly one paragraph carries algn="ctr".
    const matches = xml.match(/<a:pPr[^/]*algn="ctr"/g) ?? [];
    expect(matches.length).toBe(1);
  });

  it('setParagraphLevel writes lvl="N" for N>0 and omits it for 0', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'item 1\nnested 1\nnested 2\nitem 2',
    });
    setParagraphLevel(tb, 1, 1);
    setParagraphLevel(tb, 2, 2);

    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('lvl="1"');
    expect(xml).toContain('lvl="2"');

    // Setting back to 0 should remove the attr.
    setParagraphLevel(tb, 1, 0);
    const after = await slideXml(await savePresentation(pres), 0);
    expect((after.match(/lvl="1"/g) ?? []).length).toBe(0);
  });

  it('setParagraphLevel rejects values outside [0, 8]', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'A\nB',
    });
    expect(() => setParagraphLevel(tb, 0, 9)).toThrow(RangeError);
    expect(() => setParagraphLevel(tb, 0, -1)).toThrow(RangeError);
  });

  it('setParagraphBullet targets one paragraph with its own style', async () => {
    const { setParagraphBullet } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'A\nB\nC',
    });
    setParagraphBullet(tb, 0, 'bullet');
    setParagraphBullet(tb, 1, { char: '★' });
    setParagraphBullet(tb, 2, 'none');

    const xml = await slideXml(await savePresentation(pres), 0);
    // <a:buChar char="•"/> for paragraph 0
    expect(xml).toMatch(/<a:buChar char="•"\/?>/);
    // Star bullet for paragraph 1
    expect(xml).toContain('char="★"');
    // <a:buNone/> for paragraph 2
    expect(xml).toContain('<a:buNone');
  });

  it('setShapeText then setParagraphLevel composes cleanly', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'placeholder',
    });
    setShapeText(tb, 'parent\nchild');
    setParagraphLevel(tb, 1, 1);

    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:t>parent</a:t>');
    expect(xml).toContain('<a:t>child</a:t>');
    expect(xml).toContain('lvl="1"');
  });
});
