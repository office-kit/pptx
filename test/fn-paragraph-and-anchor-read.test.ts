// Read-back for per-paragraph alignment/level + shape text-anchor/margins.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getParagraphAlignment,
  getParagraphLevel,
  getShapeTextAnchor,
  getShapeTextMargins,
  getSlides,
  inches,
  loadPresentation,
  setParagraphAlignment,
  setParagraphLevel,
  setShapeTextAnchor,
  setShapeTextMargins,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getParagraphAlignment / getParagraphLevel', () => {
  it('round-trips alignment and level on one paragraph', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A\nB',
    });
    expect(getParagraphAlignment(tb, 0)).toBeNull();
    expect(getParagraphLevel(tb, 0)).toBe(0);

    setParagraphAlignment(tb, 1, 'center');
    setParagraphLevel(tb, 1, 2);
    expect(getParagraphAlignment(tb, 1)).toBe('ctr');
    expect(getParagraphLevel(tb, 1)).toBe(2);
  });
});

describe('fn API: getShapeTextAnchor / getShapeTextMargins', () => {
  it('round-trips anchor', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    expect(getShapeTextAnchor(tb)).toBeNull();
    setShapeTextAnchor(tb, 'top');
    expect(getShapeTextAnchor(tb)).toBe('top');
    setShapeTextAnchor(tb, 'center');
    expect(getShapeTextAnchor(tb)).toBe('center');
    setShapeTextAnchor(tb, 'bottom');
    expect(getShapeTextAnchor(tb)).toBe('bottom');
  });

  it('round-trips margins (per side)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    setShapeTextMargins(tb, { left: 1000, right: 2000 });
    const m = getShapeTextMargins(tb);
    expect(m).not.toBeNull();
    expect(m!.left).toBe(1000);
    expect(m!.right).toBe(2000);
    // Top/bottom weren't set → null.
    expect(m!.top).toBeNull();
    expect(m!.bottom).toBeNull();
  });
});
