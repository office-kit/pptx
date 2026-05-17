// getParagraphBullet — read back the bullet style on a single paragraph.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getParagraphBullet,
  getSlides,
  inches,
  loadPresentation,
  setParagraphBullet,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getParagraphBullet', () => {
  it('returns null when no bullet is set on the paragraph', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'plain',
    });
    expect(getParagraphBullet(tb, 0)).toBeNull();
  });

  it('round-trips bullet / number / none / custom char / autoNum', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A\nB\nC\nD\nE',
    });
    setParagraphBullet(tb, 0, 'bullet');
    setParagraphBullet(tb, 1, 'number');
    setParagraphBullet(tb, 2, 'none');
    setParagraphBullet(tb, 3, { char: '★' });
    setParagraphBullet(tb, 4, { autoNum: 'romanLcPeriod' });

    expect(getParagraphBullet(tb, 0)).toBe('bullet');
    expect(getParagraphBullet(tb, 1)).toBe('number');
    expect(getParagraphBullet(tb, 2)).toBe('none');
    expect(getParagraphBullet(tb, 3)).toEqual({ char: '★' });
    expect(getParagraphBullet(tb, 4)).toEqual({ autoNum: 'romanLcPeriod' });
  });
});
