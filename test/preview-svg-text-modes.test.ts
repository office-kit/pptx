// SVG-mode parity for vertical text and multi-column bodies (W1).
//
// These exercise the pure-SVG text path (`textLayout: 'svg'`, the server /
// fidelity renderer) against in-memory decks built with the public
// setShapeTextDirection / setShapeTextColumns writers, and assert it now
// honours the same layout decisions the browser (foreignObject) path expresses
// with CSS writing-mode / column-count. Import pattern follows
// test/preview-render-svg.test.ts.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  setShapeTextAutoFit,
  setShapeTextColumns,
  setShapeTextDirection,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { attrsOf, countTags } from './lib/svg-query.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  return { pres, slide: addSlide(pres, { layout }) };
};

// Enough text to wrap to several lines so column overflow / line stacking is visible.
const LONG = Array.from({ length: 24 }, (_, i) => `Line ${i + 1}`).join('\n');

// X coordinate of every body <text> (those carry text-anchor; bullets do not).
const textXs = (svg: string): number[] =>
  attrsOf(svg, 'text')
    .filter((a) => a['text-anchor'] !== undefined)
    .map((a) => Number(a.x));

describe('renderSlideToSvg — vertical text (svg mode)', () => {
  it('vert: emitted text is wrapped in a clockwise rotate() transform', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(3),
      text: LONG,
    });
    setShapeTextDirection(box, 'vert');
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toContain('transform="rotate(90');
    expect(countTags(svg, 'foreignObject')).toBe(0);
    expect(countTags(svg, 'text')).toBeGreaterThan(0);
  });

  it('vert270: rotates the opposite way (270°)', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(3),
      text: LONG,
    });
    setShapeTextDirection(box, 'vert270');
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toContain('transform="rotate(270');
    expect(svg).not.toContain('transform="rotate(90');
  });

  it('foreignObject mode still expresses vert as CSS writing-mode', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(3),
      text: LONG,
    });
    setShapeTextDirection(box, 'vert');
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
    expect(svg).toContain('writing-mode:vertical-rl');
    expect(countTags(svg, 'foreignObject')).toBeGreaterThan(0);
  });
});

describe('renderSlideToSvg — multi-column text (svg mode)', () => {
  it('2 columns: wrapped lines split across two distinct x-offsets (sequential fill)', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1.5),
      text: LONG,
    });
    setShapeTextColumns(box, { count: 2, gapEmu: 228600 }); // 0.25in gap
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const xs = textXs(svg);
    const distinct = [...new Set(xs)].sort((a, b) => a - b);
    expect(distinct.length).toBeGreaterThanOrEqual(2);
    // Column 2 sits to the right of column 1 by (colWidth + gap).
    expect(distinct[1]!).toBeGreaterThan(distinct[0]!);
  });

  it('foreignObject mode still expresses numCol as CSS column-count', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1.5),
      text: LONG,
    });
    setShapeTextColumns(box, { count: 2, gapEmu: 228600 });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
    expect(svg).toContain('column-count:2');
    expect(svg).toContain('column-gap:');
  });
});

describe('renderSlideToSvg — normAutofit shrink parity (svg vs foreignObject)', () => {
  // Largest font-size emitted, regardless of unit syntax: `font-size="N"`
  // (svg <text>) or `font-size:Npx` (foreignObject inline style).
  const maxFontPx = (svg: string): number => {
    const sizes: number[] = [];
    for (const m of svg.matchAll(/font-size(?:="|:)\s*([\d.]+)/g)) sizes.push(Number(m[1]));
    return sizes.length ? Math.max(...sizes) : 0;
  };

  // A bare <a:normAutofit/> (no baked fontScale) must shrink overflowing text to
  // fit its box — and BOTH render paths must shrink by the SAME factor, or the
  // server (svg) and browser (foreignObject) previews disagree. Normalising the
  // small-box font against a large-box baseline cancels the px-unit difference
  // between the two paths, leaving only the autofit scale.
  const shrinkRatio = async (mode: 'svg' | 'foreignObject'): Promise<number> => {
    const big = await blankSlide();
    const bigBox = addSlideTextBox(big.slide, {
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(5),
      text: LONG,
    });
    setShapeTextAutoFit(bigBox, 'normal');
    const bigFont = maxFontPx(renderSlideToSvg(big.pres, big.slide, { textLayout: mode }));

    const small = await blankSlide();
    const smallBox = addSlideTextBox(small.slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: LONG,
    });
    setShapeTextAutoFit(smallBox, 'normal');
    const smallFont = maxFontPx(renderSlideToSvg(small.pres, small.slide, { textLayout: mode }));
    return smallFont / bigFont;
  };

  it('both paths shrink a bare normAutofit body, by the same factor', async () => {
    const svgRatio = await shrinkRatio('svg');
    const foRatio = await shrinkRatio('foreignObject');
    // The overflowing small box must visibly shrink in EACH path (this is what
    // regressed: the foreignObject path used to leave bare normAutofit at 1.0).
    expect(svgRatio).toBeLessThan(0.9);
    expect(foRatio).toBeLessThan(0.9);
    // …and shrink by the same scale, so server and browser previews agree.
    expect(Math.abs(svgRatio - foRatio)).toBeLessThan(0.03);
  });
});

describe('renderSlideToSvg — horizontal parity (svg mode)', () => {
  it('a plain horizontal text box is unaffected by the W1 change (snapshot guard)', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'plain horizontal text body',
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).not.toContain('transform="rotate(');
    // The body lays out as a single left-anchored column.
    expect(new Set(textXs(svg)).size).toBe(1);
  });
});
