// setAllSlidesBackground — bulk-apply a solid background color.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideBackground,
  getSlides,
  loadPresentation,
  setAllSlidesBackground,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setAllSlidesBackground', () => {
  it('paints every slide with the requested color', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setAllSlidesBackground(pres, '#00BBCC');
    for (const slide of getSlides(pres)) {
      const bg = getSlideBackground(slide);
      expect(bg.kind).toBe('solid');
      if (bg.kind === 'solid') expect(bg.color).toBe('#00BBCC');
    }
  });
});
