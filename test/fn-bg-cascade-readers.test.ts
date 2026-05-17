// Background cascade readers — verify the layout / master variants of
// the three bg fill kinds (gradient / pattern / image) return `null`
// for layouts / masters that don't author the corresponding fill.
// Renderer cascade tests in the playground layer use real fixtures;
// these cover the API-level null contract.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideLayoutBackgroundGradientFill,
  getSlideLayoutBackgroundImageBytes,
  getSlideLayoutBackgroundPatternFill,
  getSlideMasterBackground,
  getSlideMasterBackgroundGradientFill,
  getSlideMasterBackgroundImageBytes,
  getSlideMasterBackgroundPatternFill,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: bg cascade readers null contract', () => {
  it('layout-level variants return null for unauthored fills', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      const layout = getSlideLayout(slide);
      if (!layout) continue;
      // base reader returns 'inherit' (kind) when nothing is set
      const layoutBg = getSlideLayoutBackground(layout);
      expect(layoutBg.kind).toBe('inherit');
      // each typed variant returns null on the same layout
      expect(getSlideLayoutBackgroundGradientFill(layout)).toBeNull();
      expect(getSlideLayoutBackgroundPatternFill(pres, layout)).toBeNull();
      expect(getSlideLayoutBackgroundImageBytes(pres, layout)).toBeNull();
    }
  });

  it('master-level typed variants return null when the master has a solid bg', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      const layout = getSlideLayout(slide);
      if (!layout) continue;
      // The minimal fixture's master authors a solid scheme color
      // (`{kind:'solid', color:'scheme:bg1'}`) — gradient / pattern /
      // image-fill readers must still return null.
      const masterBg = getSlideMasterBackground(pres, layout);
      expect(['solid', 'inherit']).toContain(masterBg.kind);
      expect(getSlideMasterBackgroundGradientFill(pres, layout)).toBeNull();
      expect(getSlideMasterBackgroundPatternFill(pres, layout)).toBeNull();
      expect(getSlideMasterBackgroundImageBytes(pres, layout)).toBeNull();
    }
  });
});
