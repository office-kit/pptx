// Preview rendering of preset geometry, drawn from the ECMA-376 definitions
// rather than approximations.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  inches,
  setShapeCustomGeometry,
  setShapeFill,
} from '@office-kit/pptx';
import { SHAPE_PRESETS } from '../src/internal/enum-values.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const render = (preset: string) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset,
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
  });
  setShapeFill(shape, '#4472C4');
  return { pres, slide, shape, svg: () => renderSlideToSvg(pres, slide, { textLayout: 'svg' }) };
};

describe('preset geometry in the preview', () => {
  it.each(SHAPE_PRESETS)('draws %s from its definition', (preset) => {
    const svg = render(preset).svg();
    expect(svg).not.toContain('data-pptx-preset');
    expect(svg).not.toMatch(/NaN|Infinity/);
  });

  it('shades the faces of 3-D presets over their fill', () => {
    const svg = render('cube').svg();
    expect(svg).toContain('fill="#000" fill-opacity="0.2"');
    expect(svg).toContain('fill="#fff" fill-opacity="0.2"');
  });

  it('reads arcTo angles as the direction from the center on an ellipse', () => {
    const { shape, svg } = render('rect');
    const wR = inches(2);
    const hR = inches(1);
    setShapeCustomGeometry(shape, {
      paths: [
        {
          w: inches(4),
          h: inches(2),
          fill: 'none',
          commands: [
            { kind: 'moveTo', pt: { x: wR * 2, y: hR } },
            { kind: 'arcTo', wR, hR, stAng: 0, swAng: 2_700_000 },
          ],
        },
      ],
    });
    // A 45° ray from the center meets the ellipse where x and y offsets are
    // equal: wR·hR / √(wR² + hR²) from the center. (The parametric angle would
    // land at (wR·cos 45°, hR·sin 45°) instead.)
    const offset = (wR * hR) / Math.hypot(wR, hR);
    const px = (value: number) => (value / 9525).toFixed(2);
    const end = `${px(inches(1) + wR + offset)},${px(inches(1) + hR + offset)}`;
    expect(svg()).toMatch(new RegExp(`<path d="M[^"]* ${end}"`));
  });
});
