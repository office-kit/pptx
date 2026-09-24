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
      expect(svg).toContain('<stop offset="0.0000" stop-color="#0000FF"');
      expect(svg).toContain('<stop offset="1.0000" stop-color="#FF0000"');
    },
  );
});
