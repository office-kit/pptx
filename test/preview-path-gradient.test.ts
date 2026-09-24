import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  inches,
  loadPresentation,
  setShapeGradientFill,
} from '@office-kit/pptx';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

describe('preview path gradient focus', () => {
  it.each([
    ['bottom right', 1, 1, 0, 0, '1.0000', '1.0000'],
    ['bottom left', 0, 1, 1, 0, '0.0000', '1.0000'],
    ['top right', 1, 0, 0, 1, '1.0000', '0.0000'],
    ['top left', 0, 0, 1, 1, '0.0000', '0.0000'],
    ['center', 0.5, 0.5, 0.5, 0.5, '0.5000', '0.5000'],
  ] as const)(
    'places the %s focus using edge insets',
    async (_, left, top, right, bottom, cx, cy) => {
      const pres = await loadPresentation(
        await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
      );
      const layout = findSlideLayout(pres, 'Blank');
      if (!layout) throw new Error('Blank layout missing');
      const slide = addSlide(pres, { layout });
      const shape = addSlideShape(slide, {
        preset: 'rect',
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
      });
      setShapeGradientFill(shape, {
        path: 'circle',
        focus: { left, top, right, bottom },
        stops: [
          { offset: 0, color: '#FF0000' },
          { offset: 1, color: '#0000FF' },
        ],
      });
      const svg = await renderSlideToSvg(pres, slide, { textLayout: 'svg' });
      expect(svg).toContain(`cx="${cx}" cy="${cy}"`);
      expect(svg).toContain('<stop offset="0.0000" stop-color="#FF0000"');
      expect(svg).toContain('<stop offset="1.0000" stop-color="#0000FF"');
    },
  );
});

describe('preview rectangular gradient contours', () => {
  it.each([
    ['center', 0.5, 0.5, 0.5, 0.5, 0, 0, 0, 0, [192, 192], [192, 144]],
    ['bottom right', 1, 1, 0, 0, 0, 0, -1, -1, [288, 240], [288, 192]],
    ['bottom left', 0, 1, 1, 0, -1, 0, 0, -1, [288, 240], [288, 192]],
    ['top right', 1, 0, 0, 1, 0, -1, -1, 0, [288, 144], [288, 192]],
    ['top left', 0, 0, 1, 1, -1, -1, 0, 0, [288, 144], [288, 192]],
  ] as const)(
    'keeps the %s contour rectangular',
    async (_, left, top, right, bottom, tl, tt, tr, tb, first, second) => {
      const pres = await loadPresentation(
        await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
      );
      const layout = findSlideLayout(pres, 'Blank');
      if (!layout) throw new Error('Blank layout missing');
      const slide = addSlide(pres, { layout });
      const shape = addSlideShape(slide, {
        preset: 'rect',
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
      });
      setShapeGradientFill(shape, {
        path: 'rect',
        focus: { left, top, right, bottom },
        tileRect: { left: tl, top: tt, right: tr, bottom: tb },
        stops: [
          { offset: 0, color: '#FF0000' },
          { offset: 1, color: '#0000FF' },
        ],
      });
      const { image } = renderSlideToRgba(pres, slide, { width: 960 });
      const pixel = ([x, y]: readonly [number, number]) =>
        image.data.slice((y * image.width + x) * 4, (y * image.width + x) * 4 + 3);
      const a = pixel(first),
        b = pixel(second);
      const focusPixel = pixel([Math.min(479, 96 + left * 384), Math.min(287, 96 + top * 192)]);
      expect(focusPixel[0]).toBeGreaterThan(245);
      expect(focusPixel[2]).toBeLessThan(10);
      for (let channel = 0; channel < 3; channel++)
        expect(Math.abs(a[channel]! - b[channel]!)).toBeLessThan(4);
      expect(a[0]).toBeGreaterThan(120);
      expect(a[0]).toBeLessThan(135);
      expect(a[2]).toBeGreaterThan(120);
      expect(a[2]).toBeLessThan(135);
    },
  );
});

describe('preview rectangular gradient transparency', () => {
  it('keeps opacity uniform across shared edges and a nonzero focus rectangle', async () => {
    const pres = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
    );
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('Blank layout missing');
    const slide = addSlide(pres, { layout });
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
    });
    setShapeGradientFill(shape, {
      path: 'rect',
      focus: { left: 0.4, top: 0.4, right: 0.4, bottom: 0.4 },
      stops: [
        { offset: 0, color: '#FF0000', opacity: 0.5 },
        { offset: 0.5, color: '#00FF00', opacity: 0.5 },
        { offset: 1, color: '#0000FF', opacity: 0.5 },
      ],
    });
    const { image } = renderSlideToRgba(pres, slide, { width: 960 });
    const rgb = (x: number, y: number) =>
      image.data.slice((y * image.width + x) * 4, (y * image.width + x) * 4 + 3);
    // Interior focus, an intermediate stop, and the diagonal joining two sides.
    for (const [x, y, expected] of [
      [288, 192, [255, 128, 128]],
      [173, 192, [128, 255, 128]],
      [173, 135, [128, 255, 128]],
    ] as const) {
      const actual = rgb(x, y);
      for (let channel = 0; channel < 3; channel++)
        expect(Math.abs(actual[channel]! - expected[channel]!)).toBeLessThan(5);
    }
  });
});
