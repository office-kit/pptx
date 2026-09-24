// `shapeAutoFitScale` is the shrink factor the preview applies to a
// `<a:normAutofit/>` body, exposed on its own so an editing surface painting
// its own text can shrink by exactly what the renderer did.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addTitleSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  setShapeTextAutoFit,
  type PresentationData,
} from '@office-kit/pptx';
import { renderSlideToSvg, shapeAutoFitScale } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

// Eight lines at the default size: comfortable in a 8×5 in box, far past a
// 2×1 in one, so the same text answers both halves of every case below.
const LINES = Array.from({ length: 8 }, (_, i) => `Line ${i + 1}`).join('\n');

const deck = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  return { pres, slide: addSlide(pres, { layout }) };
};

const boxWith = async (opts: { w: number; h: number; autofit: 'none' | 'normal' | 'shape' }) => {
  const { pres, slide } = await deck();
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(opts.w),
    h: inches(opts.h),
    text: LINES,
  });
  setShapeTextAutoFit(shape, opts.autofit);
  return { pres, shape };
};

describe('shapeAutoFitScale', () => {
  it('uses inherited autofit and columns in preview and editing measurements', async () => {
    const parts = unzipSync(await readFile(fixturePath));
    const layout = 'ppt/slideLayouts/slideLayout1.xml';
    parts[layout] = strToU8(
      strFromU8(parts[layout]!).replaceAll(
        '<a:bodyPr/>',
        '<a:bodyPr numCol="2" spcCol="190500"><a:normAutofit fontScale="75000"/></a:bodyPr>',
      ),
    );
    const pres = await loadPresentation(zipSync(parts));
    const slide = addTitleSlide(pres, 'Inherited title');
    const shape = getSlideShapes(slide)[0]!;
    expect(shapeAutoFitScale(pres, shape)).toBe(0.75);
    const html = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
    expect(html).toContain('column-count:2;column-gap:20.00px');
    setShapeTextAutoFit(shape, 'none');
    expect(shapeAutoFitScale(pres, shape)).toBe(1);
  });
  it('shrinks text that overflows a normAutofit box', async () => {
    const { pres, shape } = await boxWith({ w: 2, h: 1, autofit: 'normal' });
    const scale = shapeAutoFitScale(pres, shape);
    expect(scale).toBeLessThan(0.9);
    expect(scale).toBeGreaterThan(0);
  });

  it('leaves text that already fits alone', async () => {
    const { pres, shape } = await boxWith({ w: 8, h: 5, autofit: 'normal' });
    expect(shapeAutoFitScale(pres, shape)).toBe(1);
  });

  it('never shrinks a body that does not autofit', async () => {
    // noAutofit overflows and spAutoFit grows the box instead — PowerPoint
    // shrinks neither, however far the text runs past the box.
    for (const autofit of ['none', 'shape'] as const) {
      const { pres, shape } = await boxWith({ w: 2, h: 1, autofit });
      expect(shapeAutoFitScale(pres, shape)).toBe(1);
    }
  });

  it('fits into the box it is given, not the one the shape carries', async () => {
    // The editing surface asks about the box on screen, which during a resize
    // is not yet the model's.
    const { pres, shape } = await boxWith({ w: 8, h: 5, autofit: 'normal' });
    const bounds = { x: 0, y: 0, w: inches(2) as number, h: inches(1) as number };
    expect(shapeAutoFitScale(pres, shape, { bounds })).toBeLessThan(0.9);
  });

  it('agrees with what the renderer painted', async () => {
    // Same factor, or the glyphs change size the moment a caret appears. The
    // large box renders unshrunk, so the ratio between the two is the scale.
    const maxFontPx = (svg: string): number =>
      Math.max(...[...svg.matchAll(/font-size(?:="|:)\s*([\d.]+)/g)].map((m) => Number(m[1])));
    const small = await boxWith({ w: 2, h: 1, autofit: 'normal' });
    const big = await boxWith({ w: 8, h: 5, autofit: 'normal' });
    const painted = (deck: { pres: PresentationData }): number =>
      maxFontPx(
        renderSlideToSvg(deck.pres, getSlides(deck.pres)[0]!, { textLayout: 'foreignObject' }),
      );
    expect(painted(small) / painted(big)).toBeCloseTo(
      shapeAutoFitScale(small.pres, small.shape),
      5,
    );
  });
});
