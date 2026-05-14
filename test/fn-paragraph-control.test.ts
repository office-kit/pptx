// Per-paragraph alignment + nesting level.
//
// Pairs with the run-level helpers to give callers full control over a
// shape's text body without rewriting it from scratch.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  addSlideTextBox,
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
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const part = pkg.parts.find((p) => p.name === `/ppt/slides/slide${slideIndex + 1}.xml`);
  if (!part) throw new Error(`slide${slideIndex + 1}.xml not found`);
  return new TextDecoder().decode(part.data);
};

describe('fn API: per-paragraph control', () => {
  it('setParagraphAlignment targets one paragraph only', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
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
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
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
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      text: 'A\nB',
    });
    expect(() => setParagraphLevel(tb, 0, 9)).toThrow(RangeError);
    expect(() => setParagraphLevel(tb, 0, -1)).toThrow(RangeError);
  });

  it('setShapeText then setParagraphLevel composes cleanly', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
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
