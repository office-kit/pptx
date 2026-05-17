// Extended run-format attributes: strike, spc (character spacing),
// kern (kerning threshold), baseline (super / sub), cap, highlight.
// These are part of ECMA-376's CT_TextCharacterProperties surface
// (§17.18.83) and round-trip through `setShapeRunFormat` /
// `getShapeRunFormat`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeRunFormat,
  getSlides,
  inches,
  loadPresentation,
  setShapeRunFormat,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: extended run-format properties', () => {
  it('round-trips strike, spc, kern, baseline, cap, highlight', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      text: 'extended',
    });
    setShapeRunFormat(tb, 0, 0, {
      strike: true,
      spc: 200,
      kern: 1200,
      baseline: 0.3,
      cap: 'all',
      highlight: '#FFFF00',
    });
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).not.toBeNull();
    expect(fmt!.strike).toBe(true);
    expect(fmt!.spc).toBe(200);
    expect(fmt!.kern).toBe(1200);
    expect(fmt!.baseline).toBeCloseTo(0.3, 4);
    expect(fmt!.cap).toBe('all');
    expect(fmt!.highlight).toBe('#FFFF00');
  });

  it('strike accepts both boolean shorthand and explicit dblStrike', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 's',
    });
    setShapeRunFormat(tb, 0, 0, { strike: true });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe(true);
    setShapeRunFormat(tb, 0, 0, { strike: 'dblStrike' });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe('dblStrike');
    setShapeRunFormat(tb, 0, 0, { strike: false });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe(false);
  });
});
