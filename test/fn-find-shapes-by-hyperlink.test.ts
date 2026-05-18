// `findShapesByHyperlink(slide, url)` — slide-scoped finder pairing
// the presentation-level `findSlidesByHyperlink`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapesByHyperlink,
  getSlides,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesByHyperlink', () => {
  it('matches by exact-substring url', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    setShapeHyperlink(a, 'https://example.com/a');
    const b = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'B',
    });
    setShapeHyperlink(b, 'https://example.com/b');
    const c = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'C',
    });
    setShapeHyperlink(c, 'https://other.example/page');

    const matches = findShapesByHyperlink(slide, 'example.com');
    expect(matches).toHaveLength(2);
  });

  it('matches by RegExp', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    setShapeHyperlink(a, 'https://example.com/path?q=1');
    const matches = findShapesByHyperlink(slide, /[?&]q=/);
    expect(matches).toHaveLength(1);
  });

  it('returns an empty array when no shape links match', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(findShapesByHyperlink(slide, 'https://anywhere.example/')).toEqual([]);
  });
});
