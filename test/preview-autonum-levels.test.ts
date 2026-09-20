// Auto-numbering keeps one counter per indent level, like PowerPoint: a
// nested list between two top-level items must not restart the outer list.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  setParagraphBullet,
  setParagraphLevel,
  setShapeText,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const labelsIn = (svg: string): string[] =>
  [...svg.matchAll(/>\s*((?:\d+|[a-z]+)\.)\s*</g)].map((m) => m[1]!);

describe('renderSlideToSvg: auto-numbering per level', () => {
  it('continues the outer list across a nested numbered list', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(4),
      text: 'one\nnested a\nnested b\ntwo\nthree',
    });
    setShapeText(box, 'one\nnested a\nnested b\ntwo\nthree', { bullets: 'number' });
    setParagraphLevel(box, 1, 1);
    setParagraphLevel(box, 2, 1);
    setParagraphBullet(box, 1, { autoNum: 'alphaLcPeriod' });
    setParagraphBullet(box, 2, { autoNum: 'alphaLcPeriod' });

    const svg = renderSlideToSvg(pres, slide);
    expect(labelsIn(svg)).toEqual(['1.', 'a.', 'b.', '2.', '3.']);
  });

  it('restarts a level after a non-numbered paragraph at that level', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(4),
      text: 'one\ntwo\nplain\nthree',
    });
    setShapeText(box, 'one\ntwo\nplain\nthree', { bullets: 'number' });
    setParagraphBullet(box, 2, 'none');

    const svg = renderSlideToSvg(pres, slide);
    expect(labelsIn(svg)).toEqual(['1.', '2.', '1.']);
  });
});
